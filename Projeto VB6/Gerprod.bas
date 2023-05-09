Attribute VB_Name = "Gerprod"
Public UN As String
Public Familia As String

Public QT_Entrada_Estoque As Long

Public QT_Final As Boolean
Public StrSQL As String

'======================================================================
Public TotalAprovado As Double
Public TotalNaoconforme As Double
Public TotalCondicional As Double
Public Status As String
Public OrdemControlada As Boolean
Public PrimeiraOS As Boolean
Public UltimaOS As Boolean
Public Ordem As Long
Public StatusTexto As String
Public NovoValor1 As Double
Public qtdeliberada_PC As Long
Public QtdeSaida_PC As Long
Public mskdata As Date
Public LATexto As String
Public ControlaEstoque As Boolean
Public OS_texto As String


Public SQL As String
Public NumeroSerie As String
Public Codigo As Long
Public Saida As Long
Public Saida_PC As Long
Public Valor_total As Long

Public CNPJEmpresa As String
Public Individual As Boolean

Public OrdemRastreavel As Boolean

'Variaveis para criar local do bd e do relatório
Public NomeServidor         As String 'OK
Public Nome_banco           As String 'OK
Public TipoBD               As String 'OK
Public Localrel             As String 'OK
Public Usuario_banco        As String 'OK
Public Senha_banco          As String 'OK
Public LocalAntigoCaprind   As Variant 'OK
Public LocalNovoCaprind     As Variant 'OK
Public LocalAntigoGerprod   As Variant 'OK
Public LocalNovoGerprod     As Variant 'OK
Public Var                  As String 'OK
Public VarE                 As String 'OK
Public VarT                 As String 'OK
Public VarR                 As String 'OK
Public VarU                 As String 'OK
Public VarS                 As String 'OK
Public VarLAC               As String 'OK
Public VarLNC               As String 'OK
Public VarLAG               As String 'OK
Public VarLNG               As String 'OK

Public NomeServidor1        As String 'OK
Public Nome_banco1          As String 'OK
Public TipoBD1              As String 'OK
Public Localrel1            As String 'OK
Public Usuario_banco1       As String 'OK
Public Senha_banco1         As String 'OK
Public LocalAntigoCaprind1  As Variant 'OK
Public LocalNovoCaprind1    As Variant 'OK
Public LocalAntigoGerprod1  As Variant 'OK
Public LocalNovoGerprod1    As Variant 'OK
Public Var1                 As String 'OK
Public VarE1                As String 'OK
Public VarT1                As String 'OK
Public VarR1                As String 'OK
Public VarU1                As String 'OK
Public VarS1                As String 'OK
Public VarLAC1              As String 'OK
Public VarLNC1              As String 'OK
Public VarLAG1              As String 'OK
Public VarLNG1              As String 'OK

Public NomeServidor2        As String 'OK
Public Nome_banco2          As String 'OK
Public TipoBD2              As String 'OK
Public Localrel2            As String 'OK
Public Usuario_banco2       As String 'OK
Public Senha_banco2         As String 'OK
Public LocalAntigoCaprind2  As Variant 'OK
Public LocalNovoCaprind2    As Variant 'OK
Public LocalAntigoGerprod2  As Variant 'OK
Public LocalNovoGerprod2    As Variant 'OK
Public Var2                 As String 'OK
Public VarE2                As String 'OK
Public VarT2                As String 'OK
Public VarR2                As String 'OK
Public VarU2                As String 'OK
Public VarS2                As String 'OK
Public VarLAC2              As String 'OK
Public VarLNC2              As String 'OK
Public VarLAG2              As String 'OK
Public VarLNG2              As String 'OK

Public NomeServidor3        As String 'OK
Public Nome_banco3          As String 'OK
Public TipoBD3              As String 'OK
Public Localrel3            As String 'OK
Public Usuario_banco3       As String 'OK
Public Senha_banco3         As String 'OK
Public LocalAntigoCaprind3  As Variant 'OK
Public LocalNovoCaprind3    As Variant 'OK
Public LocalAntigoGerprod3  As Variant 'OK
Public LocalNovoGerprod3    As Variant 'OK
Public Var3                 As String 'OK
Public VarE3                As String 'OK
Public VarT3                As String 'OK
Public VarR3                As String 'OK
Public VarU3                As String 'OK
Public VarS3                As String 'OK
Public VarLAC3              As String 'OK
Public VarLNC3              As String 'OK
Public VarLAG3              As String 'OK
Public VarLNG3              As String 'OK

Public VerifNumero          As Variant 'Verifica campo número - OK

Public Caminho As String 'OK
Public CaminhoAnt As String 'OK
Public CaminhoNovo As String 'OK
Public FormatoData As String 'OK
Public FormatoHora As String 'OK
Public Simbolos As String 'OK
Public NomeTabelaAp         As String 'OK
Public NomeTabelaApTotalizacao     As String 'OK
Public quantnovo As Double
Public SaqueValorTotal As Double
Public IDpedido As Long
Public IDLista As Long

'======================================================================
Option Explicit
Global Salvarrel As String

Global Conexao As adodb.Connection
Public ConexaoMySql         As adodb.Connection
Public Conexao_Configuracao As adodb.Connection
Public TextoFiltro As String


Public Turno As Integer
Public TurnoMaq As Integer
Public PubUsuario As String

Public Maquina As String 'VARIAVEL DE MANUTENÇÃO
Public Controla As Boolean 'VARIAVEL DE MANUTENÇÃO
Public IDApontamento As Double 'VARIAVEL DE MANUTENÇÃO

Global Texto As String
Global BD As adodb.Connection ' Banco de dados
' Global Conexao As
Global meuwork As Workspace 'Espaço de trabalho para Base de dados

Public TBApontamento As adodb.Recordset
Public TBPlano As adodb.Recordset
Public TBAfericao As adodb.Recordset
Public TBControleNF As adodb.Recordset
Public TBOS As adodb.Recordset
Public TBOrdemServico As adodb.Recordset
Public TBProcessos As adodb.Recordset
Public TBCiclo As adodb.Recordset
Public TBProcessosDet As adodb.Recordset
Public TBUsuarios As adodb.Recordset
Public TBFases As adodb.Recordset
Public TBMaquinas As adodb.Recordset
Public TBProducao As adodb.Recordset
Public TBCodigoDesc As adodb.Recordset
Public TBOrdem As adodb.Recordset
Public TBAbrir As adodb.Recordset
Public TBGravar As adodb.Recordset
Public TBLista As adodb.Recordset
Public TBCFOP As adodb.Recordset
Public TBItem As adodb.Recordset
Public TBMaterial As adodb.Recordset
Public TBMateriaprima As adodb.Recordset
Public TBEstoque As adodb.Recordset
Public TBProduto As adodb.Recordset
Public TBFiltro As adodb.Recordset
Public TBTempo As adodb.Recordset
Public TBFI As adodb.Recordset
Public TBLogon As adodb.Recordset
Public TBMySQL As adodb.Recordset
Public TBSubreport As adodb.Recordset
Public TBUN As adodb.Recordset
Public TBUN1 As adodb.Recordset
Public TBUN2 As adodb.Recordset
Public TBVendas As adodb.Recordset

Global Situacao 'Situacao da ordem
Global Leitor As Boolean 'Uso de leitor código de barras ( sim / não )
Public ExcluiSel As Boolean 'Seleção para exclusao
Public ExcluirAP As Boolean
Public Permitido As Boolean
Public Varias_OS As Boolean 'Controla se esta apontando várias OS's ao mesmo tempo
Public Codigo_Barras As Boolean 'verifica se é apont. com código de barras
Public Ap_codigo As Boolean 'Verifica onde vai dar o setfocus para apontamento
Public TempoPreparacaoReaprov As Boolean 'Verifica se o tempo de preparacao foi reaproveitado
Public Ap_plano As Boolean 'Verifica se está sendo apontado o plano de apontamento
Public TemInternet As Boolean 'OK
Public ErroDriverMYSQL As Boolean 'OK

Public OF As Long 'Identificador da ordem
Public OS As Long 'Identificado da ordem de serviço
Public IDProcesso As Long 'Identificador do processo de fabricação
Public IDFase As Long ' Identificador da fase do processo
Public IDProducao As Long ' Identificador da ordem de serviço
Public i As Long 'OK

Public Quant As Long ' Quantidade de peças a produzir
Public QuantEmpenhoPC As Double 'OK
Public Produzidas As Double 'Quantidade de peças produzidas
Public QtdeSaida As Double 'Quantidade de peças NC - OK
Public Entrada   As Double 'OK
Public Valor_CSLL_Serv As Double
Public Valor_IPI As Double
Public Valor_CSLL_Prod As Double
Public Valor_Cofins_Prod As Double
Public Valor_Cofins_Serv As Double
Public Valor_INSS_Serv As Double
Public QuantComprado As Double
Public QuantComprado1 As Double
Public Qtd_Prog As Double
Public ValorNC As Double
Public ValorConta As Double
Public Qtde As Double
Public qt As Double
Public qtdeliberar As Double
Public Valorhora As Double
Public ValorhoraPrep As Double
Public SaqueSaldo As Double
Public QtdeRefugo As Double

Public Evento As String 'Código do evento
Public DescEvento As String 'Descrição do evento

Public Gravar As Boolean 'Gravar Sim/não
Global OrdemExiste As Boolean 'Verifica se existe ordem de servico

'Variaveis de atualização da ordem de servico
Public QTOK As Double 'Quantidade de pecas OK
Public QTNC As Double 'Quantidade de pecas não conforme
Public QTCD As Double 'Quantidade de pecas condicional

Public TPPREV As Date 'Tempo de preparação previsto
Public TPPSEG As Double 'Tempo de preparação previsto em segundos
Public TTPUTIL As Variant 'Tempo total de preparação previsto
Public TPUTIL As Date 'Tempo de preparação utilizado
Public TPUSEG As Double 'Tempo de preparação utilizado em segundos
Public TPUSEGDECS As Double 'Tempo de preparação utilizado em segundos - OK

Public TEPREV As Date 'Tempo de execução previsto
Public TEPSEG As Double 'Tempo de execução previsto em segundos
Public TEUTIL As Date 'Tempo de execução utilizado
Public TEUSEG As Double 'Tempo de execução Utilizado em segundos
Public DecimoSegundos   As Double 'OK

Public TempoTotalUtil As Variant ' Somatorio dos total de execução e preparação
Public TTUTILSEG As Double 'Tempo total Utilizado na OS em segundos (execução + preparação)
Public TTUTIL As Variant 'Tempo total utilizado na OS
Public TTEUTILS As Double 'Tempo total utilizado na OS em segundos

Public Hora_apontamento As Date 'Data e hora do apontamento

Public CMSEG As Double 'Custo máquina em reais
Public CTTLOTE As Double 'Custo total do lote
Public CTTPECA As Double 'Custo total da peça
Public QuantEmpenho As Double

Public TTNC As Double 'Total de peças não conforme
Public TTOK As Double 'Total de peças conforme
Public TTCD As Double 'Total de peças condicional

Public LOTE As Double 'Total de peças a ser produzidas

Public Contador As Integer
Public Contador2 As Integer
Public Qtlicencas_gerprod As Integer
Public Eficiencia_prep As Double
Public Eficiencia_exec As Double

'Centro de custo
Public Id_item As Integer
Public codproduto As Long
Public IDAntigo As Long

'Variaveis para entrada e baixa do estoque
Public Valortotal As Double
Public quantestoque As Double
Public qtdeliberada As Double
Public QtdeEstoque As Double
Public Qtd As Double
Public Dimensoes As Double

'Variaveis de fechamento de quantidades da ordem
Public TOK As Double 'Total aprovadas
Public TNC As Double 'Total Não conforme
Public TCD As Double 'Total aprovada condicional

Public IDLogon As Long
Public pubRegistrado As String
Public pubLicenca As String
Public Operador As String
'Public Descricao As String
'Public Fase As String
Public NomeCampo As Variant 'Controle de mensagens - OK
Public Acao As String 'Controle de mensagens - OK
Public Versao_processo As String 'Versão do processo

Public Ultimo, Penultimo As Integer 'Variavel de identificação do ultimo e penultimo evento

'Localizar pasta
Public Const BIF_RETURNONLYFSDIRS = 1 'OK
Public Const BIF_DONTGOBELOWDOMAIN = 2 'OK
Public Const MAX_PATH = 260 'OK

Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long 'OK
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long 'OK
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long 'OK

Public Type BrowseInfo 'OK
    hWndOwner As Long 'OK
    pIDLRoot As Long 'OK
    pszDisplayName As Long 'OK
    lpszTitle As Long 'OK
    ulFlags As Long 'OK
    lpfnCallback As Long 'OK
    lParam As Long 'OK
    iImage As Long 'OK
End Type

Public lpIDList As Long 'OK
Public sBuffer As String 'OK
Public szTitle As String 'OK
Public tBrowseInfo As BrowseInfo 'OK

Public iFlow As Integer, iTempEcho As Boolean 'OK

'Visualizar arquivos
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Const SW_SHOWDEFAULT = 10

'Relatórios crystal 11
Public crAPP As New CRAXDDRT.Application
Public Report As CRAXDDRT.Report
Public crxExport As CRAXDDRT.ExportOptions
Public CPProperty As CRAXDDRT.ConnectionProperty
Public DBTable As CRAXDDRT.DatabaseTable
Public SubReport As CRAXDDRT.SubreportObject
Public Nomerel As String
Public FormulaRel As String
Public NomeSubReport As String, NomeSubReport1 As String, NomeSubReport2 As String, NomeSubReport3 As String, NomeSubReport4 As String, NomeSubReport5 As String 'OK
Public NomeSubReport6 As String, NomeSubReport7 As String, NomeSubReport8 As String, NomeSubReport9 As String, NomeSubReport10 'OK
Public PermitidoRel As Boolean
Public LocalRelPersonalizado As String
Public LocalrelNovo As String

'Resolução da tela/Monitor
Public xTwips%, yTwips%, xPixels#, YPixels# 'OK
Public xPixelsAnt As Long, YPixelsAnt As Long 'OK

'Atualizando o sistema
Public LocalAntigoSincCaprind As Variant
Public LocalNovoSincCaprind As Variant
Public LocalAntigoSincGerprod As Variant
Public LocalNovoSincGerprod As Variant
Global Fso, f, fG, Fsu, FU, FUG
Global Caprind As String
Global Gerprod As String
Public Atualizando As Boolean

Public Familiatext As String 'OK
Public NovoValor As String 'OK
Public INNERJOINTEXTO As String 'OK
Public MensagemErro As String

'Carrega Instância do SQL automaticamente
Const SEPARATOR        As String = ""
Public vSplit          As Variant
Public sSrv            As String
Public sDb             As String
Public sUser           As String
Public sPass           As String
Public vSrv            As Variant
Public vDb             As Variant
Public sText           As String
Public m_bEnumSrv      As Boolean
Public m_bEnumDbOdbc   As Boolean
Public m_bEnumDbAdo    As Boolean
Public m_bOk           As Boolean

'Criar aquivos (.txt), gerenciar pastas e arquivos
Public GerArqPastas As New FileSystemObject 'OK
Public ArqTXT As TextStream 'OK
Public Arq As Long
Public DadosArquivo As String
Public Linha As String

'Acesso a internet
Public Ie As Object 'Chat online OK

Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Const CONNECT_LAN As Long = &H2
Private Const CONNECT_MODEM As Long = &H1
Private Const CONNECT_PROXY As Long = &H4
Private Const CONNECT_RAS As Long = &H10
Private Const CONNECT_OFFLINE As Long = &H20
Private Const CONNECT_CONFIGURED As Long = &H40

'Alterar resolução do monitor
Public Const AW_HOR_POSITIVE = &H1 'Animates the window from left to right. This flag can be used with roll or slide animation.
Public Const AW_HOR_NEGATIVE = &H2 'Animates the window from right to left. This flag can be used with roll or slide animation.
Public Const AW_VER_POSITIVE = &H4 'Animates the window from top to bottom. This flag can be used with roll or slide animation.
Public Const AW_VER_NEGATIVE = &H8 'Animates the window from bottom to top. This flag can be used with roll or slide animation.
Public Const MF_CHECKED = &H8&
Public Const MF_APPEND = &H100&
Public Const TPM_LEFTALIGN = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const BITSPIXEL = 12
Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public nDC As Long

'Verificar focu no programa
Declare Function CallFunWindowProc Lib "user32" Alias "CallFunWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const WM_ACTIVATEAPP = &H1C
Public Const GWL_WNDPROC = -4
Global lpPrevWndProc As Long
Global gHW As Long

'Verifica hora do servidor
Private Declare Function NetRemoteTOD Lib "NETAPI32.DLL" (ByVal server As String, buffer As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function NetApiBufferFree Lib "NETAPI32.DLL" (buffer As Any) As Long

Private Type TIME_OF_DAY
  t_elapsedt As Long
  t_msecs As Long
  t_hours As Long
  t_mins As Long
  t_secs As Long
  t_hunds As Long
  t_timezone As Long
  t_tinterval As Long
  t_day As Long
  t_month As Long
  t_year As Long
  t_weekday As Long
End Type

Dim DatainiTexto As String 'OK
Public ArrayQtdeDescNC() As String 'OK

' Função nova simular enter
Public Sub Sendkeys(Text$, Optional wait As Boolean = False)
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.Sendkeys Text, wait
                    Set WshShell = Nothing
End Sub

Function FunKeyAscii(KA As Integer)
On Error GoTo tratar_erro

If KA = 13 Then
    KA = 0
    Sendkeys "{TAB}"
End If
1:
FunKeyAscii = KA

Exit Function
tratar_erro:
    If Err.Number = "70" Then
        GoTo 1:
        Exit Function
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcRemoveObjetosResize(Formulario As Form)
On Error GoTo tratar_erro
Dim Controle As Control

For i = 0 To Formulario.Controls.Count - 1
    Set Controle = Formulario.Controls(i)
    If TypeOf Controle Is ListView Then Controle.Tag = "Ex_Columns,Ex_Font"
    If TypeOf Controle Is MSFlexGrid Then Controle.Tag = "Ex_Font"
    If TypeOf Controle Is USToolBar Then Controle.Tag = "Ex_Font"
    If TypeOf Controle Is USTab Then Controle.Tag = "Ex_Font"
    
'    If TypeOf Controle Is USTreeView Then Controle.Tag = "Ex_Font"
    
'    If TypeOf Controle Is Label Then Controle.Tag = "Ex_Font"
'    If TypeOf Controle Is TextBox Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is ComboBox Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is USButton Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is MaskEdBox Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is frame Then Controle.Tag = "Ex_Font"
'    If TypeOf Controle Is Button Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is DTPicker Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is SSTab Then Controle.Tag = "Ex_Font"
'    If TypeOf Controle Is CommandButton Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is Image Then Controle.Tag = "Ex_Height"
'    If TypeOf Controle Is USTextBox Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is OptionButton Then Controle.Tag = "Ex_Height, Ex_Font"
'    If TypeOf Controle Is CheckBox Then Controle.Tag = "Ex_Height, Ex_Font"
Next i


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function FunAbreBD() As Boolean
On Error GoTo tratar_erro
    
Set Conexao = New adodb.Connection
With Conexao
    .Provider = "SQLOLEDB"
    .Properties("Data Source").Value = NomeServidor
    .Properties("Initial catalog").Value = Nome_banco
    .Properties("User ID").Value = IIf(Usuario_banco = "", "Procam", Usuario_banco)
    .Properties("Password").Value = IIf(Senha_banco = "", "PRO0902loc$?", Senha_banco)
    .Properties("Persist Security Info") = "False"
    .Open
End With
FunAbreBD = True

Exit Function
tratar_erro:
    If Err.Number = "-2147467259" Then
        FunAbreBD = False
        If MsgBox("Não foi possivel acessar o banco de dados, deseja configurar antes de utilizar o sistema?", vbYesNo) = vbNo Then End
        frmOpcoesGeral2.Show 1
        Exit Function
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function FunAbreBDSite() As Boolean
On Error GoTo tratar_erro

'MYSQL Site
Permitido = True
Atualizando = False
ErroDriverMYSQL = False
Set ConexaoMySql = New adodb.Connection
With ConexaoMySql
    'Porta 3306
    .ConnectionTimeout = 60
    .CommandTimeout = 400
    .CursorLocation = adUseClient

Conectar:
    If Permitido = False Then
        .Open "DRIVER={MySQL ODBC 3.51 Driver};" & "user=caprind11" & ";password=cap0902loc" & ";database=caprind11" & ";server=mysql02.caprind1.hospedagemdesites.ws" & ";option=20499"
    Else
        .Open "DRIVER={MySQL ODBC 5.1 Driver};" & "user=caprind11" & ";password=cap0902loc" & ";database=caprind11" & ";server=mysql02.caprind1.hospedagemdesites.ws" & ";option=20499"
'        Caminho = App.Path & "\AtualizacaoDriverMYSQL5.1.13.txt"
'        Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
'        If GerArqPastas.FileExists(Caminho) = False Then
'            Arq = FreeFile 'Esta linha atribui um espaço livre na memória para armazenar o arquivo.
'            Open Caminho For Append As #Arq 'Escreve em um arquivo já existente
'            DadosArquivo = Now & " - Driver MySql atualizado para versão 5.1.13"
'            Print #Arq, DadosArquivo; 'Escreve a data limite para utilização sem internet
'            Close (Arq) 'Fecha o arquivo
'            GerArqPastas.GetFile(Caminho).Attributes = Hidden
'        End If
    End If
End With

Exit Function
tratar_erro:
    If Err.Number = "-2147467259" Then
        'Caminho = App.Path & "\AtualizacaoDriverMYSQL5.1.13.txt"
        'Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
        'If GerArqPastas.FileExists(Caminho) = False Then
        
        Caminho = "C:\Program Files (x86)\MySQL\Connector ODBC 5.1"
        Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
        If GerArqPastas.FolderExists(Caminho) = False Then
            Caminho = "C:\Program Files\MySQL\Connector ODBC 5.1"
            Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
            If GerArqPastas.FolderExists(Caminho) = False Then
                Call DS.USMsgBox("É obrigatório atualizar o driver MySQL antes de logar." & vbCrLf & "IMPORTANTE: clique no botão Next> da instalação do driver para prosseguir." & vbCrLf & "O sistema será encerrado após a atualização.", vbInformation, "Driver MySQL desatualizado")
                Atualizando = True
                With frmfundo.kftp
                    .DisableRESTCommand
                    If FunConectaKFTP(frmfundo.kftp) = False Then
                        Permitido = False
                        Atualizando = False
                        GoTo Conectar
                    End If
                End With
                If FunDownloadKFTP(frmfundo.kftp, "mysql-connector-odbc-5.1.13-win32.msi", App.Path & "\mysql-connector-odbc-5.1.13-win32.msi") = False Then
                    Permitido = False
                    Atualizando = False
                    GoTo Conectar
                End If
                ProcAbrirArquivo (App.Path & "\mysql-connector-odbc-5.1.13-win32.msi")
                End
            End If
        End If
        ErroDriverMYSQL = True
        Exit Function
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function FunFechaBDSite()
On Error GoTo tratar_erro

If ConexaoMySql.State = 1 Then ConexaoMySql.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Sub ProcValidaSenha()
On Error GoTo tratar_erro

'*************************************************
'* ROTINA DE LIBERAÇÃO DO CAPRIND PARA UTILIZAÇÃO
'*************************************************

'Dim licensa As LicensaTFB
'Dim EXISTE
'Dim serial As serialTFB

'EXISTE = Dir("C:\Procam\licensa.dat", vbDirectory + vbHidden)
'If EXISTE = "" Then
'MsgBox ("Execute o administrador para liberar a utilização do Gerprod 2002 v3.01."), vbInformation
'End
'End If

'Open "C:\PROCAM\licensa.dat" For Random As #1 Len = 70
'Get #1, 1, licensa

'Open "C:\arq.log" For Random As #2 Len = 300

'Get #2, 1, serial
'If licensa.serie = serial.contador Then

'If licensa.contador = licensa.numero Then
'        MsgBox ("ESTE APLICATIVO VENCEU O NÚMERO DE EXECUÇÕES LIBERADAS PARA VERSÃO DEMO, ENTRE EM CONTATO COM A PROCAM - PROGRAMAÇÃO C.N.C - TEL : 0XX19-3894-8046."), vbCritical, "Caprind 2002 - Procam programação c.n.c - Tel : 0xx19-3894-8046"
'    Close #1
'    End
'End If

'If licensa.contador <> 9999 Then
'        MsgBox ("ÉSTA É UMA VERSÃO DEMO DO GERPROD 2002 V3.01 QUE ESTÁ LIBERADA PARA EXECUTAR " & licensa.numero & " VEZ(ES), E JÁ EXECUTOU " & licensa.contador & " VEZ(ES), FALTA(M) " & licensa.numero - licensa.contador & " VEZ(ES) PARA VENCER A VERSÃO DEMO, SOLICITE A LIBERAÇÃO TOTAL JUNTO A PROCAM."), vbExclamation, "Caprind 2002 - Procam programação c.n.c - Tel : 0xx19-3894-8005"
'        licensa.contador = licensa.contador + 1
'    Put #1, 1, licensa
'    Close #1
'End If

'If licensa.contador = 9999 Then
'        MsgBox ("ÉSTA É UMA VERSÃO LIBERADA DO GERPROD 2002 V3.01 JUNTO A PROCAM."), vbExclamation, "Caprind 2002 - Procam programação c.n.c - Tel : 0xx19-3894-8005"
'    Close #1
'End If

'*************************************************
'End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub Main()
On Error GoTo tratar_erro
Dim VersaoNova As Integer 'OK

'Verifica local de execução do exe
Caminho = App.Path
Contador = Len(Caminho)
If Right(Caminho, 16) <> "Projetos\Gerprod" And Mid(Caminho, 4, 21) <> "Arquivos de Programas" And Mid(Caminho, 4, 21) <> "Arquivos de programas" And Mid(Caminho, 4, 27) <> "Arquivos de Programas (x86)" And Mid(Caminho, 4, 27) <> "Arquivos de programas (x86)" And Mid(Caminho, 4, 13) <> "Program Files" And Mid(Caminho, 4, 13) <> "Program files" And Mid(Caminho, 4, 19) <> "Program Files (x86)" And Mid(Caminho, 4, 19) <> "Program files (x86)" Then
    'MsgBox ("Não é permitido abrir o Gerprod deste caminho " & Caminho & "."), vbCritical
    'End
End If

'Verifica resolução do computador
xTwips = Screen.TwipsPerPixelX
yTwips = Screen.TwipsPerPixelY
xPixels = Screen.Width / xTwips
YPixels = Screen.Height / yTwips
xPixelsAnt = xPixels
YPixelsAnt = YPixels

ProcCarregaBancoDados
ProcVerifAtualizacao

FormatoData = GetSetting("Procam", "CaprindSQL", "FormatoData", "dd/mm/yyyy")
FormatoHora = GetSetting("Procam", "CaprindSQL", "FormatoHora", "hh:mm:ss")
Simbolos = "Ø±¼½¾²³ª°¡¢£¤¥¦§¨©­®¯´µ·¸¹º¿æ÷øð×"
pubLicenca = "Máximo 10 liberações"
pubRegistrado = "Demonstração"
If Salvarrel = "" Then Salvarrel = False
'Se o banco de dados for localizado.
If Salvarrel = False Then
    With frmabertura
        .Timer1.Enabled = True
        .Timer2.Enabled = True
        .Show
    End With
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerifAtualizacao()
On Error GoTo tratar_erro

LocalAntigoGerprod = GetSetting("Procam", "CaprindSQL", "LocalAntigoGerprod")
LocalNovoGerprod = GetSetting("Procam", "CaprindSQL", "LocalNovoGerprod")

If LocalAntigoGerprod <> "" Then
    CaminhoAnt = Left((LocalAntigoGerprod), Len(LocalAntigoGerprod) - 12)
End If

If LocalNovoGerprod <> "" Then
    CaminhoNovo = Left(LocalNovoGerprod, Len(LocalNovoGerprod) - 12)
End If

If LocalAntigoGerprod <> "" And LocalNovoGerprod <> "" Then
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set Fsu = CreateObject("Scripting.FileSystemObject")
    NomeCampo = "Gerprod.exe na pasta " & CaminhoAnt
    Set f = Fso.GetFile(LocalAntigoGerprod)
    NomeCampo = "Gerprod.exe na pasta " & CaminhoNovo
    Set FU = Fsu.GetFile(LocalNovoGerprod)
    If f.DateLastModified < FU.DateLastModified Then
    NomeCampo = "SincGerprod.exe na pasta " & CaminhoNovo
    MsgBox ("O sistema Gerprod está desatualizado e será atualizado automaticamente."), vbInformation
    Shell CaminhoNovo & "\SincGerprod.exe", vbNormalFocus
    End
    End If
End If
1:

Exit Sub
tratar_erro:
    If Err.Number = "53" Then
        MsgBox ("Não será possivel atualizar o sistema, pois não foi encontrado o arquivo " & NomeCampo & "."), vbExclamation
        GoTo 1
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaBancoDados()
On Error GoTo tratar_erro

NomeServidor = GetSetting("Procam", "CaprindSQL", "NomeServidor")
TipoBD = GetSetting("Procam", "CaprindSQL", "TipoBD")
Localrel = GetSetting("Procam", "CaprindSQL", "Localrel")
Nome_banco = GetSetting("Procam", "CaprindSQL", "Nome_banco")
Usuario_banco = GetSetting("Procam", "CaprindSQL", "Usuario_banco")
Senha_banco = GetSetting("Procam", "CaprindSQL", "Senha_banco")
LocalAntigoCaprind = GetSetting("Procam", "CaprindSQL", "LocalAntigoCaprind")
LocalNovoCaprind = GetSetting("Procam", "CaprindSQL", "LocalNovoCaprind")
LocalAntigoGerprod = GetSetting("Procam", "CaprindSQL", "LocalAntigoGerprod")
LocalNovoGerprod = GetSetting("Procam", "CaprindSQL", "LocalNovoGerprod")

NomeServidor1 = GetSetting("Procam", "CaprindSQL", "NomeServidor1")
TipoBD1 = GetSetting("Procam", "CaprindSQL", "TipoBD1")
Localrel1 = GetSetting("Procam", "CaprindSQL", "Localrel1")
Nome_banco1 = GetSetting("Procam", "CaprindSQL", "Nome_banco1")
Usuario_banco1 = GetSetting("Procam", "CaprindSQL", "Usuario_banco1")
Senha_banco1 = GetSetting("Procam", "CaprindSQL", "Senha_banco1")
LocalAntigoCaprind1 = GetSetting("Procam", "CaprindSQL", "LocalAntigoCaprind1")
LocalNovoCaprind1 = GetSetting("Procam", "CaprindSQL", "LocalNovoCaprind1")
LocalAntigoGerprod1 = GetSetting("Procam", "CaprindSQL", "LocalAntigoGerprod1")
LocalNovoGerprod1 = GetSetting("Procam", "CaprindSQL", "LocalNovoGerprod1")

NomeServidor2 = GetSetting("Procam", "CaprindSQL", "NomeServidor2")
TipoBD2 = GetSetting("Procam", "CaprindSQL", "TipoBD2")
Localrel2 = GetSetting("Procam", "CaprindSQL", "Localrel2")
Nome_banco2 = GetSetting("Procam", "CaprindSQL", "Nome_banco2")
Usuario_banco2 = GetSetting("Procam", "CaprindSQL", "Usuario_banco2")
Senha_banco2 = GetSetting("Procam", "CaprindSQL", "Senha_banco2")
LocalAntigoCaprind2 = GetSetting("Procam", "CaprindSQL", "LocalAntigoCaprind2")
LocalNovoCaprind2 = GetSetting("Procam", "CaprindSQL", "LocalNovoCaprind2")
LocalAntigoGerprod2 = GetSetting("Procam", "CaprindSQL", "LocalAntigoGerprod2")
LocalNovoGerprod2 = GetSetting("Procam", "CaprindSQL", "LocalNovoGerprod2")

NomeServidor3 = GetSetting("Procam", "CaprindSQL", "NomeServidor3")
TipoBD3 = GetSetting("Procam", "CaprindSQL", "TipoBD3")
Localrel3 = GetSetting("Procam", "CaprindSQL", "Localrel3")
Nome_banco3 = GetSetting("Procam", "CaprindSQL", "Nome_banco3")
Usuario_banco3 = GetSetting("Procam", "CaprindSQL", "Usuario_banco3")
Senha_banco3 = GetSetting("Procam", "CaprindSQL", "Senha_banco3")
LocalAntigoCaprind3 = GetSetting("Procam", "CaprindSQL", "LocalAntigoCaprind3")
LocalNovoCaprind3 = GetSetting("Procam", "CaprindSQL", "LocalNovoCaprind3")
LocalAntigoGerprod3 = GetSetting("Procam", "CaprindSQL", "LocalAntigoGerprod3")
LocalNovoGerprod3 = GetSetting("Procam", "CaprindSQL", "LocalNovoGerprod3")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunLogonIn(Usuario As String) As Boolean
On Error GoTo tratar_erro

FunLogonIn = True
Set TBLogon = CreateObject("adodb.recordset")
If Usuario <> "PROCAM" Then
    TBLogon.Open "Select * from Logon where usuario = '" & Usuario & "' and Tipo = 'G'", Conexao, adOpenKeyset, adLockOptimistic
    If TBLogon.EOF = False Then
        If MsgBox("O usuário " & Usuario & " já está conectado, deseja desconectar a conexão antiga e iniciar uma nova conexão?", vbYesNo) = vbYes Then
            Conexao.Execute "DELETE from Logon WHERE usuario = '" & Usuario & "' and Tipo = 'G'"
            
            'Conta 11 segundos para desconectar a outra conexão
            Dataini = Format(Now, "hh:mm:ss")
            Dataini = Dataini + "00:00:11"
            Do While Format(Now, "hh:mm:ss") < Dataini
            
            Loop
            
            Set TBLogon = CreateObject("adodb.recordset")
            TBLogon.Open "Select * from Logon", Conexao, adOpenKeyset, adLockOptimistic
            TBLogon.AddNew
        Else
            FunLogonIn = False
        End If
    Else
        TBLogon.AddNew
    End If
Else
    Conexao.Execute "DELETE from Logon WHERE usuario = '" & Usuario & "' and Tipo = 'G'"
    TBLogon.Open "Select * from Logon", Conexao, adOpenKeyset, adLockOptimistic
    TBLogon.AddNew
End If
TBLogon!Usuario = Usuario
TBLogon!Data = Date
TBLogon!Hora = Time
'TBLogon!Hora_ultimo_evento = Time
TBLogon!Tipo = "G"
TBLogon.Update
IDLogon = TBLogon!IDLogon
TBLogon.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Sub ProcLogonOut(Usuario As String)
On Error GoTo tratar_erro

Conexao.Execute "DELETE from Logon where usuario = '" & Usuario & "' and Tipo = 'G'"

'Efetua logoff do usuário no site
If TemInternet = True And ErroDriverMYSQL = False Then
    Set TBFiltro = CreateObject("adodb.recordset")
    TBFiltro.Open "Select * from Empresa where CNPJ = '" & CNPJEmpresa & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFiltro.EOF = False Then
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then ConexaoMySql.Execute "Update usuarios Set Logado_gerprod = 'NÃO' where CNPJ = '" & TBFiltro!CNPJ & "' and Usuario = '" & Usuario & "'"
    End If
    TBFiltro.Close
End If
                
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

                               'ORDEM       QTDE. PREVISTA              QTDE. OK      QT. PROD.(OK+NC) CUSTO LOTE       CUSTO PEÇA          CUSTO TERCEIROS         CUSTO MATERIAL         CUSTO OUTRAS          ORDEM CONSIGNADA
Function FunCalculaValorUnitOrdem(OF As Long, QuantsolicitadoN1 As Double, qt As Double, Qtde As Double, CTLote As Double, CTPecaReal As Double, CTTerceiros As Double, CTMaterial As Double, CTOutras As Double, Consignada As Boolean)
On Error GoTo tratar_erro
Dim TBMaterialVlrUnitOrdem As adodb.Recordset
Dim TBEstoqueVlrUnitOrdem As adodb.Recordset
Dim TBproducaoVlrUnitOrdem As adodb.Recordset
Dim TBAbrirVlrUnitOrdem As adodb.Recordset
Dim Permitido_Calula_Vlr_Unit_Ordem As Boolean

'Valor NC
Valor_Cofins_Prod = 0
Valor1 = 0 'Serviço
Valor2 = 0 'Material
Valor3 = 0 'Mão de obra
ValorConta = 0 'Outras
Valor_CSLL_Serv = 0
Valor_INSS_Serv = 0
Valor_IPI = 0
Permitido_Calula_Vlr_Unit_Ordem = False

'Custo de material
If Consignada = False Then
    Set TBMaterialVlrUnitOrdem = CreateObject("adodb.recordset")
    TBMaterialVlrUnitOrdem.Open "Select Valor_saida_estoque, Saida from Producaomaterial where Ordem = " & OF & " order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaterialVlrUnitOrdem.EOF = False Then
        Do While TBMaterialVlrUnitOrdem.EOF = False
            
            'Verifica valor total do material
            If TBMaterialVlrUnitOrdem!Valor_saida_estoque > 0 Then
            Valor_CSLL_Prod = IIf(IsNull(TBMaterialVlrUnitOrdem!Valor_saida_estoque), 0, TBMaterialVlrUnitOrdem!Valor_saida_estoque)
            Else
            Valor_CSLL_Prod = 0
            End If
            
            If TBMaterialVlrUnitOrdem!Saida <> "NÃO" Then
                Set TBproducaoVlrUnitOrdem = CreateObject("adodb.recordset")
                TBproducaoVlrUnitOrdem.Open "Select Totalprod from ordemservico where Ordem = " & OF & " ORDER BY fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                If TBproducaoVlrUnitOrdem.EOF = False Then
                    Qtd_Prog = IIf(IsNull(TBproducaoVlrUnitOrdem!Totalprod), 0, TBproducaoVlrUnitOrdem!Totalprod) 'Qtde. produzida
                    If Qtd_Prog <> 0 And Valor_CSLL_Prod <> 0 Then
                        Valor_CSLL_Serv = Format(Valor_CSLL_Serv + (Valor_CSLL_Prod / Qtd_Prog), "###,##0.0000000000")
                    ElseIf Qtde <> 0 And Valor_CSLL_Prod <> 0 Then
                            Valor_CSLL_Serv = Format(Valor_CSLL_Serv + (Valor_CSLL_Prod / Qtde), "###,##0.0000000000")
                    End If
                End If
                TBproducaoVlrUnitOrdem.Close
            End If
            TBMaterialVlrUnitOrdem.MoveNext
        Loop
    End If
    TBMaterialVlrUnitOrdem.Close
End If

'Verifica qtde NC da ordem
QuantComprado = 0
Set TBAbrirVlrUnitOrdem = CreateObject("adodb.recordset")
TBAbrirVlrUnitOrdem.Open "Select Sum(TTNC) as QtdeNC from CQ_NC_FABRICA where Ordem = " & OF & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrirVlrUnitOrdem.EOF = False Then
    QuantComprado = IIf(IsNull(TBAbrirVlrUnitOrdem!QtdeNC), 0, TBAbrirVlrUnitOrdem!QtdeNC)
End If

'Verifica última OS com NC
Set TBAbrirVlrUnitOrdem = CreateObject("adodb.recordset")
TBAbrirVlrUnitOrdem.Open "Select OS.Fase FROM ordemservico OS INNER JOIN CQ_NC_FABRICA CQNC ON OS.Idproducao = CQNC.OS where OS.Ordem = " & OF & " and CQNC.PARECERCQ = 'Rejeitar' order by OS.fase, OS.retrabalho, OS.IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrirVlrUnitOrdem.EOF = False Then
    TBAbrirVlrUnitOrdem.MoveLast
    OS = TBAbrirVlrUnitOrdem!Fase
End If
Set TBproducaoVlrUnitOrdem = CreateObject("adodb.recordset")
TBproducaoVlrUnitOrdem.Open "Select * from ordemservico where Ordem = " & OF & " and Fase <= " & OS & " ORDER BY fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBproducaoVlrUnitOrdem.EOF = False Then
    Do While TBproducaoVlrUnitOrdem.EOF = False
        'Soma valor unitário do SERVIÇO na OS
        If IsNull(TBproducaoVlrUnitOrdem!Totalprod) = False And TBproducaoVlrUnitOrdem!Totalprod <> "" And TBproducaoVlrUnitOrdem!Totalprod <> "0" Then Valor_IPI = Format(Valor_IPI + (TBproducaoVlrUnitOrdem!CTServico / TBproducaoVlrUnitOrdem!Totalprod), "###,##0.0000000000")
        
        'Soma valor unitário da MÃO DE OBRA na OS
        Valor_Cofins_Prod = Format(Valor_Cofins_Prod + TBproducaoVlrUnitOrdem!CRPECA, "###,##0.0000000000")
        'Verifica qtde. regufada na OS
        Qtd_Prog = 0
        Set TBAbrirVlrUnitOrdem = CreateObject("adodb.recordset")
        TBAbrirVlrUnitOrdem.Open "Select Sum(TTNC) as QtdeNC from CQ_NC_FABRICA where OS = " & TBproducaoVlrUnitOrdem!IDProducao & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrirVlrUnitOrdem.EOF = False Then
            Permitido_Calula_Vlr_Unit_Ordem = True
            Qtd_Prog = IIf(IsNull(TBAbrirVlrUnitOrdem!QtdeNC), 0, TBAbrirVlrUnitOrdem!QtdeNC)
            Valor1 = Format(Valor1 + (Valor_IPI * Qtd_Prog), "###,##0.00") 'Valor total unitário serviço x qtde. refugada da OS
            Valor3 = Format(Valor3 + (Valor_Cofins_Prod * Qtd_Prog), "###,##0.00") 'Valor total unitário mão de obra x qtde. refugada da OS
        End If
        TBAbrirVlrUnitOrdem.Close
        TBproducaoVlrUnitOrdem.MoveNext
    Loop
End If
If Permitido_Calula_Vlr_Unit_Ordem = True Then
    'Valor do material por peça x qtde. refugada
    If QuantsolicitadoN1 <> 0 Then Valor2 = Format(Valor_CSLL_Serv * QuantComprado, "###,##0.00")
                       'SE  +   MT   +   MO
    ValorNC = Format(Valor1 + Valor2 + Valor3, "###,##0.00")
Else
    ValorNC = 0
End If

'Custo de MO do lote
Valor_Cofins_Prod = CTLote

'Custo de terceiros por peça
Valor2 = 0
Set TBproducaoVlrUnitOrdem = CreateObject("adodb.recordset")
TBproducaoVlrUnitOrdem.Open "Select Totalprod, CTServico from ordemservico where Ordem = " & OF & " and Custos = 'False' ORDER BY fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBproducaoVlrUnitOrdem.EOF = False Then
    Do While TBproducaoVlrUnitOrdem.EOF = False
        If TBproducaoVlrUnitOrdem!Totalprod <> 0 Then
            Valor2 = Valor2 + (IIf(IsNull(TBproducaoVlrUnitOrdem!CTServico), 0, TBproducaoVlrUnitOrdem!CTServico) / IIf(IsNull(TBproducaoVlrUnitOrdem!Totalprod), 0, TBproducaoVlrUnitOrdem!Totalprod))
        ElseIf Qtde <> 0 Then
                Valor2 = Valor2 + (IIf(IsNull(TBproducaoVlrUnitOrdem!CTServico), 0, TBproducaoVlrUnitOrdem!CTServico) / Qtde)
        End If
        TBproducaoVlrUnitOrdem.MoveNext
    Loop
End If
TBproducaoVlrUnitOrdem.Close

'Valor por peça
'Custo total
Valor_Cofins_Prod = Format(CTTerceiros + CTMaterial + CTLote + CTOutras, "###,##0.00")
If ValorNC <> 0 And QuantsolicitadoN1 = Qtde Then
                                            'Custo total  - Valor NC / QTDE. OK                                                               Custo total  - Valor NC
    If qt <> 0 Then FunCalculaValorUnitOrdem = Format((Valor_Cofins_Prod - ValorNC) / qt, "###,##0.0000000000") Else FunCalculaValorUnitOrdem = Format(Valor_Cofins_Prod - ValorNC, "###,##0.0000000000")
Else
    If Qtde > 0 Then CTOutras = CTOutras / Qtde Else CTOutras = 0
                                    'SE    +        MT       +  MO        + OU
    FunCalculaValorUnitOrdem = Format((Valor2 + Valor_CSLL_Serv + CTPecaReal + CTOutras), "###,##0.0000000000")
End If

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Sub ProcVerificaAcao()
On Error GoTo tratar_erro

MsgBox ("Informe " & NomeCampo & " antes de " & Acao & "."), vbExclamation

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcEmpenhaProdutoeAtualQtdeEntEmpOrdem(LOTE As Long, Codinterno As String, Qtde_entrada As Double, ID_estoque As Long)
On Error GoTo tratar_erro

'Verifica se a ordem tem pedido vinculado, empenha o produto e atualiza a quantidade de entrada no empenho da ordem
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select VC.Qtde_produzir, VC.qtdeexpedida, VC.CODIGO, PP.Qtde_empenho, PP.Qtde_entrada from (producao_pedidos PP INNER JOIN vendas_carteira VC ON PP.IDcarteira = VC.Codigo) INNER JOIN Producao P ON P.Ordem = PP.Ordem where P.Ordem = " & LOTE & " and P.Desenho = '" & Codinterno & "' and VC.Desenho = '" & Codinterno & "' and PP.Expedicao = 'False' and VC.Cotacao <> 0 order by VC.Prazofinal", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Qtd = IIf(IsNull(TBAbrir!Qtde_produzir), 0, TBAbrir!Qtde_produzir) - IIf(IsNull(TBAbrir!qtdeexpedida), 0, TBAbrir!qtdeexpedida)
        
        'Verifica quantidade já empenhada
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Sum(Qtde_empenhada) as qtde, Sum(Qtde_saida) as Saida from Estoque_Controle_Empenho_Vendas where ID_carteira = " & TBAbrir!Codigo, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Qtde = IIf(IsNull(TBFI!Qtde), 0, TBFI!Qtde) - IIf(IsNull(TBFI!Saida), 0, TBFI!Saida)
        End If
        TBFI.Close
        
        Dimensoes = Qtd - Qtde
        If Dimensoes > 0 Then
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Estoque_Controle_Empenho_Vendas", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!Data = Date
            TBGravar!Responsavel = Operador
            TBGravar!ID_estoque = ID_estoque
            TBGravar!ID_carteira = TBAbrir!Codigo
            If Qtde_entrada >= Dimensoes Then TBGravar!Qtde_empenhada = Dimensoes Else TBGravar!Qtde_empenhada = Qtde_entrada
            TBGravar.Update
            TBGravar.Close
            Qtde_entrada = Qtde_entrada - Dimensoes
        End If
        If Qtde_entrada <= 0 Then GoTo Prosseguir
        TBAbrir.MoveNext
    Loop
End If
Prosseguir:

ProcAtualizaQtdeEntEmpProd LOTE, Codinterno

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAtualizaQtdeEntEmpProd(LOTE As Long, Codinterno As String)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select PP.* from producao_pedidos PP INNER JOIN Producao P ON P.Ordem = PP.Ordem where P.Ordem = " & LOTE & " and P.Desenho = '" & Codinterno & "' order by PP.IDcarteira", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    
    Set TBTempo = CreateObject("adodb.recordset")
    TBTempo.Open "Select Sum(Entrada) as Qtde from Estoque_movimentacao where Lote = '" & LOTE & "' and Desenho = '" & Codinterno & "' and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
    If TBTempo.EOF = False Then
        Qtde = IIf(IsNull(TBTempo!Qtde), 0, TBTempo!Qtde)
    End If
    
    Do While TBAbrir.EOF = False
        If Qtde > 0 Then
            If Qtde > TBAbrir!Qtde_empenho Then
                TBAbrir!Qtde_entrada = TBAbrir!Qtde_empenho
                Qtde = Qtde - TBAbrir!Qtde_empenho
            Else
                TBAbrir!Qtde_entrada = Qtde
                Qtde = 0
            End If
            TBAbrir.Update
        Else
            TBAbrir!Qtde_entrada = 0
        End If
        TBAbrir.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcVerifRelPersonalizado()
On Error GoTo tratar_erro

PermitidoRel = False
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select * from Qualidade_revisao_relatorios where Nome_relatorio = '" & Nomerel & "' and Personalizado = 'True' order by Revisao desc, ID desc", Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    LocalRelPersonalizado = Localrel & "\Personalizados"
    PermitidoRel = True
End If
TBMaterial.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunCalculaQtdePCKG(Qtde_est_KG As Double, Qtde_est_PC As Double, Qtde_movimentacao As Double, CalculaPC As Boolean) As Double
On Error GoTo tratar_erro
Dim Kg_un As Double 'OK

FunCalculaQtdePCKG = 0
If Qtde_est_PC <> 0 Then
    Kg_un = Format(Qtde_est_KG / Qtde_est_PC, "###,##0.0000000000")
    If CalculaPC = True Then FunCalculaQtdePCKG = Qtde_movimentacao / Kg_un Else FunCalculaQtdePCKG = Qtde_movimentacao * Kg_un
End If
    
Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcAbrirArquivo(Caminho As String)
On Error GoTo tratar_erro

Call ShellExecute(0&, vbNullString, Caminho, vbNullString, vbNullString, SW_SHOWDEFAULT)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcVerifInstancia(Combo As ComboBox)
On Error GoTo tratar_erro

With Combo
    .Clear
    Screen.MousePointer = vbHourglass
    For Each vSrv In EnumSqlServers
        .AddItem vSrv
    Next
    m_bEnumSrv = True
    Screen.MousePointer = vbDefault
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunVerifMovimentacaoEstPC(ID_empresa As Long) As Boolean
On Error GoTo tratar_erro

FunVerifMovimentacaoEstPC = False
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select * from Empresa where Codigo = " & ID_empresa & " and Movimentar_estoque_pc = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    FunVerifMovimentacaoEstPC = True
End If
TBTempo.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcVerificaNumero()
On Error GoTo tratar_erro

If (IsNumeric(VerifNumero)) = False Or InStr(VerifNumero, "-") <> 0 And InStr(VerifNumero, "-") <> 1 Then
    MsgBox ("Só é permitido número neste campo."), vbExclamation
    VerifNumero = False
Else
    VerifNumero = True
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcVerifAPCodigo(Plano_prod As Boolean, TextoFiltro As String)
On Error GoTo tratar_erro

Ap_codigo = False
If Plano_prod = True Then
    INNERJOINTEXTO = "((ProducaoFases_OS PFO INNER JOIN ordemservico OS ON OS.ID_apontamento = PFO.ID) INNER JOIN Producao P ON P.Ordem = OS.Ordem) INNER JOIN Empresa E ON E.Codigo = P.ID_empresa"
    NomeCampo = "PFO.ID"
Else
    INNERJOINTEXTO = "Producao P INNER JOIN Empresa E ON E.Codigo = P.ID_empresa"
    NomeCampo = "P.Ordem"
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select E.Codigo from " & INNERJOINTEXTO & " where " & NomeCampo & " = " & TextoFiltro & " and E.Apontamento_codigo = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Ap_codigo = True
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunValidarCliente(Usuario As String) As Boolean
On Error GoTo tratar_erro

FunValidarCliente = True
Set TBFiltro = CreateObject("adodb.recordset")
TBFiltro.Open "Select * from Empresa where CNPJ = '" & CNPJEmpresa & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFiltro.EOF = False Then
    Do While TBFiltro.EOF = False
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then
            Set TBMySQL = New adodb.Recordset
            TBMySQL.Open "Select * From Clientes Where CNPJ = '" & TBFiltro!CNPJ & "'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            With TBMySQL
                If .EOF = False Then
                    .Fields!NomeRazao = TBFiltro!Razao
                    .Update
                    
                    TBFiltro!Licencas_caprind = IIf(IsNull(.Fields!Licencas), 0, .Fields!Licencas)
                    TBFiltro!Licencas_gerprod = IIf(IsNull(.Fields!Licencas_gerprod), 0, .Fields!Licencas_gerprod)
                    TBFiltro!Modulo = .Fields!Modulo
                    TBFiltro.Update
                    
                    Permitido = True
                    If .Fields!Liberado = "NÃO" Then
                        Permitido = False
                        If IsNull(.Fields!Codigo_do_erro) = True Or .Fields!Codigo_do_erro = 0 Then
                            MensagemErro = "Não é possível efetuar o apontamento, pois o mesmo está com o acesso bloqueado"
                        Else
                            Select Case .Fields!Codigo_do_erro
                                Case 1: MensagemErro = "Error Accessing the system registry"
                                Case 2: MensagemErro = "Invalid procedure call or argument"
                                Case 3: MensagemErro = "Out of memory"
                                Case 4: MensagemErro = "Many client applications trying to access the DLL at the same time"
                                Case 5: MensagemErro = "Server object was not properly registered or not found"
                                Case 6: MensagemErro = "ActiveX Control not found"
                                Case 7: MensagemErro = "License information for this component not found or not found DLL"
                                Case 8: MensagemErro = "Invalid number of arguments or invalid property assignment"
                                Case 9: MensagemErro = "Syntax error (missing operator)"
                            End Select
                        End If
                    End If
                    If Permitido = False Then
                        MsgBox (MensagemErro & "."), vbCritical
                        FunValidarCliente = False
                        Exit Function
                    End If
                    
                    'Verifica número de licenças
                    If IsNull(.Fields!Licencas_gerprod) = False And .Fields!Licencas_gerprod <> "" And Usuario <> "PROCAM" Then
                        If .Fields!Licencas_gerprod = 0 Then
                            USMsgBox ("Não é possível efetuar o apontamento, pois o Gerprod não está incluso em contrato."), vbCritical, "CAPRIND v5.0"
                            FunValidarCliente = False
                            Exit Function
                        Else
                            SQL = "Delete from Logon Where Data < '" & Date & "' and Tipo = 'G'"
                            'Debug.Print SQL
                            
                            Conexao.Execute SQL
                            
                            Set TBLogon = CreateObject("adodb.recordset")
                            TBLogon.Open "Select * from Logon where Usuario <> 'PROCAM' and Data = '" & Date & "' and Tipo = 'G'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBLogon.EOF = False Then
                                Contador = TBLogon.RecordCount
                                Contador2 = .Fields!Licencas_gerprod
                                If Contador > Contador2 Then
                                    USMsgBox ("Atenção!" & vbCrLf & "Não é possível efetuar o apontamento, pois ja foram utilizados todas as licenças disponiveis." & vbCrLf & "Licenças Disponíveis : " & Contador2 & vbCrLf & "Licenças utilizadas : " & Contador), vbCritical, "CAPRIND v5.0 - GERPROD"
                                    FunValidarCliente = False
                                    Exit Function
                                End If
                            End If
                            TBLogon.Close
                        End If
                    End If
                End If
            End With
        End If
        TBFiltro.MoveNext
    Loop
End If
TBFiltro.Close
FunFechaBDSite

If FunValidarUsuario(Usuario) = False Then FunValidarCliente = False

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Function FunValidarUsuario(Usuario As String) As Boolean
On Error GoTo tratar_erro
Dim P As String 'OK

FunValidarUsuario = True
Set TBFiltro = CreateObject("adodb.recordset")
TBFiltro.Open "Select * from Empresa where CNPJ = '" & CNPJEmpresa & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFiltro.EOF = False Then
    Do While TBFiltro.EOF = False
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then
            Set TBMySQL = New adodb.Recordset
            TBMySQL.Open "Select * From usuarios Where CNPJ = '" & TBFiltro!CNPJ & "' and Usuario = '" & Usuario & "'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            With TBMySQL
                If .EOF = True Then
                    .AddNew
                    .Fields!Liberado = "SIM"
                Else
                    Permitido = True
                    If .Fields!Liberado = "NÃO" Then
                        Permitido = False
                        If IsNull(.Fields!Codigo_do_erro) = True Or .Fields!Codigo_do_erro = 0 Then
                            MensagemErro = "Não é possível o usuário " & Usuario & " efetuar o apontamento, pois o mesmo está com o acesso bloqueado"
                        Else
                            Select Case .Fields!Codigo_do_erro
                                Case 1: MensagemErro = "Error Accessing the system registry"
                                Case 2: MensagemErro = "Invalid procedure call or argument."
                                Case 3: MensagemErro = "Out of memory"
                                Case 4: MensagemErro = "Many client applications trying to access the DLL at the same time"
                                Case 5: MensagemErro = "Server object was not properly registered or not found"
                                Case 6: MensagemErro = "ActiveX Control not found"
                                Case 7: MensagemErro = "License information for this component not found or not found DLL"
                                Case 8: MensagemErro = "Invalid number of arguments or invalid property assignment"
                                Case 9: MensagemErro = "Syntax error (missing operator)"
                            End Select
                        End If
                    End If
                    If Permitido = False Then
                        MsgBox (MensagemErro & "."), vbCritical
                        FunValidarUsuario = False
                        Exit Function
                    End If
                End If
                .Fields!CNPJ = TBFiltro!CNPJ
                .Fields!Usuario = Usuario
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Nome, Senha, Setor, Email from Usuarios where Usuario = '" & Usuario & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Fields!Nome = TBAbrir!Nome
                    .Fields!Senha = TBAbrir!Senha
                    .Fields!Cargo = TBAbrir!Setor
                    If TBAbrir!Email <> "" Then .Fields!Email = TBAbrir!Email
                End If
                TBAbrir.Close
                
                .Fields!Logado_gerprod = "SIM"
                .Fields!Nivel = 2
                .Fields!Ativo = 1
                .Update
                
                Caminho = App.Path & "\CIGE.txt"
                Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                If GerArqPastas.FileExists(Caminho) = True Then GerArqPastas.DeleteFile (Caminho) 'Deleta o arquivo antigo
                
                Arq = FreeFile 'Esta linha atribui um espaço livre na memória para armazenar o arquivo.
                Open Caminho For Append As #Arq 'Escreve em um arquivo já existente
                
                'Verifica a data limite para utilização sem internet e converte (15 dias)
                Contador = 1
                Dataini = Date + 15
                DadosArquivo = "@%&*#"
                Do While Contador <= 8
                    Select Case Mid(Format(Dataini, "dd/mm/yy"), Contador, 1)
                        Case 0: P = "!"
                        Case 1: P = "#"
                        Case 2: P = "S"
                        Case 3: P = "&"
                        Case 4: P = "|"
                        Case 5: P = "Z"
                        Case 6: P = "@"
                        Case 7: P = "$"
                        Case 8: P = "T"
                        Case 9: P = "^"
                        Case "/": P = "?"
                    End Select
                    DadosArquivo = DadosArquivo & P
                    Contador = Contador + 1
                Loop
                DadosArquivo = DadosArquivo & "&*$@!"
                Print #Arq, DadosArquivo; 'Escreve a data limite para utilização sem internet
                Close (Arq) 'Fecha o arquivo
                GerArqPastas.GetFile(Caminho).Attributes = Hidden
                
            End With
        End If
        TBFiltro.MoveNext
    Loop
End If

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Function FunValidarClienteSemInternet() As Boolean
On Error GoTo tratar_erro

FunValidarClienteSemInternet = True
Set TBFiltro = CreateObject("adodb.recordset")
TBFiltro.Open "Select Licencas_gerprod from Empresa where CNPJ = '" & CNPJEmpresa & "' and Licencas_gerprod IS NOT NULL and Licencas_gerprod <> N'' and Licencas_gerprod > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBFiltro.EOF = False Then
    Do While TBFiltro.EOF = False
        Set TBLogon = CreateObject("adodb.recordset")
        TBLogon.Open "Select * from Logon where Usuario <> 'PROCAM' and Data = '" & Date & "' and Tipo = 'G'", Conexao, adOpenKeyset, adLockOptimistic
        If TBLogon.EOF = False Then
            Contador = TBLogon.RecordCount
            Contador2 = TBFiltro!Licencas_gerprod
            If Contador > Contador2 Then
                    USMsgBox ("Atenção!" & vbCrLf & "Não é possível efetuar o apontamento, pois ja foram utilizados todas as licenças disponiveis." & vbCrLf & "Licenças Disponíveis : " & Contador2 & vbCrLf & "Licenças utilizadas : " & Contador), vbCritical, "CAPRIND v5.0 - GERPROD"
                    FunValidarClienteSemInternet = False
                Exit Function
            End If
        End If
        TBLogon.Close
        TBFiltro.MoveNext
    Loop
Else
    USMsgBox ("Não é possível efetuar o apontamento, pois o Gerprod não está incluso em contrato."), vbCritical, "GERPROD"
    FunValidarClienteSemInternet = False
    Exit Function
End If
TBFiltro.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcVerificaInternet()
On Error GoTo tratar_erro
Dim IP As String 'OK

TemInternet = True
'If FunInternetConectada("") = False Then
If IsInternetOnline = False Then
    ProcVerifDiasUtilizadosSemInternet
    TemInternet = False
'Else
'    IP = frmabertura.Inet1.OpenURL("http://www.google.com.br/")
'    If IP = "" Then
'        ProcVerifDiasUtilizadosSemInternet
'        TemInternet = False
'    End If
End If
If TemInternet = True Then
    DatainiTexto = Date
    SaveSetting "{6F6CC9481G35-A412-3500-A1Z1-B58654D8}", "Default", "Main", "FT58SV98Q3" & FunTamanhoTextoZeroEsq(Day(Date), 2) & "5ER8ASE1" & FunTamanhoTextoZeroEsq(Month(Date), 2) & "7EBL5Q" & Year(Date) & "EFFO895Q"
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Function FunInternetConectada(Optional ByRef ConnType As String) As Boolean
On Error GoTo tratar_erro
Dim dwFlags As Long
Dim WebTest As Boolean

ConnType = ""
WebTest = InternetGetConnectedState(dwFlags, 0&)
Select Case WebTest
    Case dwFlags And CONNECT_LAN: ConnType = "LAN"
    Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
    Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
    Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
    Case dwFlags And CONNECT_CONFIGURED: ConnType = "Configurada"
    Case dwFlags And CONNECT_RAS: ConnType = "Remota"
End Select
FunInternetConectada = WebTest

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcVerifDiasUtilizadosSemInternet()
On Error GoTo tratar_erro
Dim P As String
Dim DataTexto As String
Dim DataTextoAtual As String
Dim Datafim As Date
Dim DiaTexto As String
Dim MesTexto As String
Dim AnoTexto As String

Caminho = App.Path & "\CIGE.txt"

'Verifica se o arquivo esta na pasta
Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
If GerArqPastas.FileExists(Caminho) = False Then
    MsgBox ("Ocorreu um erro inesperado, o sistema será encerrado."), vbCritical
    End
Else
    'Verifica se o arquivo não foi alterado
    Arq = FreeFile 'Esta linha atribui um espaço livre na memória para armazenar o arquivo.
    Open Caminho For Input As #Arq 'Abre o arquivo
    Line Input #Arq, Linha 'Ler uma linha
    
    Permitido = True
    If Len(Linha) <> 18 Then
        Permitido = False
        GoTo Parar
    End If
    
    'Verifica se existe algum caracter diferente
    Contador = 6
    Do While Contador <= 13
        If Mid(Linha, Contador, 1) <> "!" And Mid(Linha, Contador, 1) <> "#" And Mid(Linha, Contador, 1) <> "S" And Mid(Linha, Contador, 1) <> "&" And Mid(Linha, Contador, 1) <> "|" And Mid(Linha, Contador, 1) <> "Z" And Mid(Linha, Contador, 1) <> "@" And Mid(Linha, Contador, 1) <> "$" And Mid(Linha, Contador, 1) <> "T" And Mid(Linha, Contador, 1) <> "^" And Mid(Linha, Contador, 1) <> "?" Then Permitido = False
        Contador = Contador + 1
    Loop
    Close (Arq) 'Fecha o arquivo

Parar:
    If Permitido = False Then
        MsgBox ("Ocorreu um erro inesperado, o sistema será encerrado."), vbCritical
        End
    End If
    
    'Verifica dados do arquivo e carrega nas variaveis
    Arq = FreeFile 'Esta linha atribui um espaço livre na memória para armazenar o arquivo.
    Open Caminho For Input As #Arq 'Abre o arquivo
    Line Input #Arq, Linha 'Ler uma linha
    
    DataTexto = ""
    'Verifica a data que esta no txt e converte
    Contador = 6
    Do While Contador <= 13
        Select Case Mid(Linha, Contador, 1)
            Case "!": P = 0
            Case "#": P = 1
            Case "S": P = 2
            Case "&": P = 3
            Case "|": P = 4
            Case "Z": P = 5
            Case "@": P = 6
            Case "$": P = 7
            Case "T": P = 8
            Case "^": P = 9
            Case "?": P = "/"
        End Select
        If DataTexto = "" Then DataTexto = P Else DataTexto = DataTexto & P
        Contador = Contador + 1
    Loop
    Dataini = DataTexto
    If Format(Date, "dd/mm/yy") > Dataini Then
        MsgBox ("Venceu o prazo de 15 dias para utilização do Gerprod sem internet, o sistema será encerrado. Favor entar em contato com o suporte através do e-mail suporte@caprind.com.br."), vbCritical
        End
    End If
    Close (Arq) 'Fecha o arquivo
    
    DatainiTexto = GetSetting("{6F6CC9481G35-A412-3500-A1Z1-B58654D8}", "Default", "Main") 'Verifica a data inicial sem internet
    If DatainiTexto <> "" Then
        DatainiTexto = GetSetting("{6F6CC9481G35-A412-3500-A1Z1-B58654D8}", "Default", "Main") 'Carrega data incial sem internet
        DiaTexto = Mid(DatainiTexto, 11, 2)
        MesTexto = Mid(DatainiTexto, 21, 2)
        AnoTexto = Mid(DatainiTexto, 29, 4)
        DatainiTexto = DiaTexto & "/" & MesTexto & "/" & AnoTexto
    Else
        DatainiTexto = Date
        SaveSetting "{6F6CC9481G35-A412-3500-A1Z1-B58654D8}", "Default", "Main", "FT58SV98Q3" & FunTamanhoTextoZeroEsq(Day(Date), 2) & "5ER8ASE1" & FunTamanhoTextoZeroEsq(Month(Date), 2) & "7EBL5Q" & Year(Date) & "EFFO895Q" 'Salva data incial sem internet
    End If
    Dataini = DatainiTexto
    Dia = DateDiff("d", Dataini, Date)
    If Dia > 15 Then
        MsgBox ("Venceu o prazo de 15 dias para utilização do Gerprod sem internet, o sistema será encerrado. Favor entar em contato com o suporte através do e-mail suporte@caprind.com.br."), vbCritical
        End
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcLogonOutSemUtilizacao()
On Error GoTo tratar_erro

Set TBLogon = CreateObject("adodb.recordset")
TBLogon.Open "Select Usuario from Logon where Data < '" & Date & "' and Tipo = 'G'", Conexao, adOpenKeyset, adLockOptimistic
If TBLogon.EOF = False Then
    Do While TBLogon.EOF = False
        Conexao.Execute "DELETE from Logon WHERE usuario = '" & TBLogon!Usuario & "' and Tipo = 'G'"
        Conexao.Execute "DELETE from Producao_Relatorios where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Producao_Relatorios_Total where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Estoque_relatorios where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Troca_titulo_relatorio where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Plano_de_contas_totalizacao where Responsavel = '" & TBLogon!Usuario & "'"
        Conexao.Execute "DELETE from Compras_Recebimento_Relatorios where responsavel = '" & TBLogon!Usuario & "'"
        
        'Efetua logoff do usuário no site
        If TemInternet = True And ErroDriverMYSQL = False Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select IDLogon from Logon where usuario = '" & TBLogon!Usuario & "' and Data = '" & Date & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                Set TBFiltro = CreateObject("adodb.recordset")
                TBFiltro.Open "Select * from Empresa where CNPJ = '" & CNPJEmpresa & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFiltro.EOF = False Then
                    FunAbreBDSite
                    If ConexaoMySql.State = 1 Then ConexaoMySql.Execute "Update usuarios Set Logado = 'NÃO', Logado_Gerprod = 'NÃO' where CNPJ = '" & TBFiltro!CNPJ & "' and Usuario = '" & TBLogon!Usuario & "'"
                End If
                TBFiltro.Close
            End If
            TBAbrir.Close
        End If
        TBLogon.MoveNext
    Loop
End If
TBLogon.Close
                
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunGotFocus(Objeto)
On Error GoTo tratar_erro

Objeto.SelStart = 0
Objeto.SelLength = Len(Objeto.Text)

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcCarregaToolBar1(Formulario As Form, Comprimento As Integer, QtdeBotao As Integer, Visivel As Boolean)
On Error GoTo tratar_erro

With Formulario.USToolBar1
    .DrawButtonsEx Formulario.USImageList1
    .Height = 975
    .Width = Comprimento
    Contador = QtdeBotao
    Do While QtdeBotao > 0
        .ButtonForeColor(QtdeBotao) = &H0&
        .ButtonFont(QtdeBotao).Bold = True
        QtdeBotao = QtdeBotao - 1
    Loop
    .Refresh
    If Visivel = True Then .Visible = True Else .Visible = False
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunVerifQtdeRetrabalhoFase(Ordem As Long, Fase As String) As Double
On Error GoTo tratar_erro

FunVerifQtdeRetrabalhoFase = 0
Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select Sum(QTOK) as TTOK from ordemservico where Ordem = " & Ordem & " and Fase = '" & Fase & "' and Retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
    FunVerifQtdeRetrabalhoFase = IIf(IsNull(TBFases!TTOK), 0, TBFases!TTOK)
End If
TBFases.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcAlteraResolucaoMonitor(X As Long, Y As Long, Bits As Long)
On Error GoTo tratar_erro
Dim DevM As DEVMODE

Familiatext = EnumDisplaySettings(0&, 0&, DevM)
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
DevM.dmPelsWidth = X 'ScreenWidth
DevM.dmPelsHeight = Y 'ScreenHeight
DevM.dmBitsPerPel = Bits '(can be 8, 16, 24, 32 or even 4)
Familiatext = ChangeDisplaySettings(DevM, CDS_TEST) 'Now change the display and check if possible
If Familiatext = DISP_CHANGE_SUCCESSFUL Then
    Familiatext = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
ElseIf Familiatext <> DISP_CHANGE_RESTART Then
        MsgBox ("Resolução de vídeo não suportada, o sistema será encerrado."), vbExclamation
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

'=============VERIFICA FOCU DO PROGRAMA================
Public Sub Hook()
On Error GoTo tratar_erro

lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf FunWindowProc)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub Unhook()
On Error GoTo tratar_erro
Dim temp As Long

temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo tratar_erro
   
If uMsg = WM_ACTIVATEAPP Then
    If wParam <> 0 Then 'Com focu
        If xPixels = 800 And YPixels = 600 Or xPixels = 1024 And YPixels = 768 Then
        
        Else
            nDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
            ProcAlteraResolucaoMonitor 1024, 768, GetDeviceCaps(nDC, BITSPIXEL)
        End If
    Else 'Sem focu
        If xPixelsAnt > 1024 And YPixelsAnt = 768 Or xPixelsAnt = 1024 And YPixelsAnt > 768 Or xPixelsAnt > 1024 And YPixelsAnt > 768 Then
            ProcAlteraResolucaoMonitor xPixelsAnt, YPixelsAnt, GetDeviceCaps(nDC, BITSPIXEL)
            DeleteDC nDC
        End If
    End If
End If
FunWindowProc = CallFunWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

'=============FIM VERIFICA FOCU DO PROGRAMA================

Sub ProcEmpenharREAutomOrdem(IDestoque As Long, Qtde_entrada As Double, LOTE As String, Data As Date, Responsavel As String, Codinterno As String, Excluir As Boolean)
On Error GoTo tratar_erro

'Empenha para a ordem vinculada ao mesmo pedido interno que requisita o material
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select PP.OrdemEmpenho, PP.Qtde_empenho from Producao_pedidos PP INNER JOIN vendas_carteira VC ON VC.Codigo = PP.IDcarteira where PP.Ordem = " & LOTE & " and PP.OrdemEmpenho IS NOT NULL and PP.OrdemEmpenho <> 0 order by VC.Prazofinal", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    Do While TBCFOP.EOF = False And Qtde_entrada > 0
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Producao_NF_Consignada where Data = '" & Data & "' and Responsavel = '" & Responsavel & "' and Ordem = " & TBCFOP!OrdemEmpenho & " and IDestoque = " & IDestoque, Conexao, adOpenKeyset, adLockOptimistic
        If Excluir = False Then
            'Verifica quantidade empenhada
            Qtde = TBCFOP!Qtde_empenho
            
            If TBGravar.EOF = True Then TBGravar.AddNew
            TBGravar!Data = Data
            TBGravar!Responsavel = Responsavel
            TBGravar!Ordem = TBCFOP!OrdemEmpenho
            TBGravar!Codinterno = Codinterno
            TBGravar!IDestoque = IDestoque
            
            If Qtde_entrada > Qtde Then
                TBGravar!Quantidade = TBGravar!Quantidade + Qtde
                Qtde_entrada = Qtde_entrada - Qtde
            Else
                TBGravar!Quantidade = TBGravar!Quantidade + Qtde_entrada
                Qtde_entrada = 0
            End If
            TBGravar!Quantidade_PC = TBGravar!Quantidade
            TBGravar.Update
        Else
            If TBGravar.EOF = False Then
                If TBGravar!Quantidade - Qtde_entrada <= 0 Then
                    Qtde_entrada = Qtde_entrada - TBGravar!Quantidade
                    TBGravar.Delete
                Else
                    If Qtde_entrada > TBGravar!Quantidade Then
                        Qtde_entrada = Qtde_entrada - TBGravar!Quantidade
                        TBGravar!Quantidade = 0
                    Else
                        TBGravar!Quantidade = TBGravar!Quantidade - Qtde_entrada
                        Qtde_entrada = 0
                    End If
                    TBGravar!Quantidade_PC = TBGravar!Quantidade
                    TBGravar.Update
                End If
            End If
        End If
        TBGravar.Close
Proximo:
        TBCFOP.MoveNext
    Loop
End If
TBCFOP.Close

'Set TBCFOP = CreateObject("adodb.recordset")
'TBCFOP.Open "Select PP.OrdemEmpenho from Producao_pedidos PP INNER JOIN vendas_carteira VC ON VC.Codigo = PP.IDcarteira where PP.Ordem = " & LOTE & " and PP.OrdemEmpenho IS NOT NULL and PP.OrdemEmpenho <> 0 order by VC.Prazofinal", Conexao, adOpenKeyset, adLockOptimistic
'If TBCFOP.EOF = False Then
'    Do While TBCFOP.EOF = False And Qtde_entrada > 0
'        Set TBTempo = CreateObject("adodb.recordset")
'        If Excluir = False Then
'            TBTempo.Open "Select PM.Ordem, PM.Requisitado - (ISNULL(QSEP.Saida, 0) + ISNULL(QEPD.Qtde_empenhar, 0)) AS Requisitado from ((Producaomaterial PM INNER JOIN Producao P ON P.Ordem = PM.Ordem) LEFT JOIN Qtde_saida_estoque_produto QSEP ON QSEP.Ordem = PM.Ordem and QSEP.Desenho = PM.Codigo) LEFT JOIN Qtde_empenhada_produto_detalhado QEPD ON QEPD.Ordem = PM.Ordem and QEPD.Codigo = PM.Codigo where PM.Ordem = " & TBCFOP!OrdemEmpenho & " and PM.Codigo = '" & Codinterno & "' and PM.Saida <> 'SIM' and PM.Requisitado - (ISNULL(QSEP.Saida, 0) + ISNULL(QEPD.Qtde_empenhar, 0)) > 0 order by PM.Ordem desc", Conexao, adOpenKeyset, adLockOptimistic
'        Else
'            TBTempo.Open "Select PM.Ordem from Producaomaterial PM INNER JOIN Producao P ON P.Ordem = PM.Ordem where PM.Ordem = " & TBCFOP!OrdemEmpenho & " and PM.Codigo = '" & Codinterno & "' order by PM.Ordem desc", Conexao, adOpenKeyset, adLockOptimistic
'        End If
'        If TBTempo.EOF = False Then
'            Do While TBTempo.EOF = False And Qtde_entrada > 0
'                Set TBGravar = CreateObject("adodb.recordset")
'                TBGravar.Open "Select * from Producao_NF_Consignada where Data = '" & Data & "' and Responsavel = '" & Responsavel & "' and Ordem = " & TBTempo!Ordem & " and IDestoque = " & IDestoque, Conexao, adOpenKeyset, adLockOptimistic
'
'                If Excluir = False Then
'                    If TBGravar.EOF = True Then TBGravar.AddNew
'                    TBGravar!Data = Data
'                    TBGravar!Responsavel = Responsavel
'                    TBGravar!Ordem = TBTempo!Ordem
'                    TBGravar!Codinterno = Codinterno
'                    TBGravar!IDestoque = IDestoque
'                    If Qtde_entrada > TBTempo!Requisitado Then
'                        TBGravar!Quantidade = TBGravar!Quantidade + TBTempo!Requisitado
'                        Qtde_entrada = Qtde_entrada - TBTempo!Requisitado
'                    Else
'                        TBGravar!Quantidade = TBGravar!Quantidade + Qtde_entrada
'                        Qtde_entrada = 0
'                    End If
'                    TBGravar!Quantidade_PC = TBGravar!Quantidade
'                    TBGravar.Update
'                Else
'                    If TBGravar.EOF = False Then
'                        If TBGravar!Quantidade - Qtde_entrada <= 0 Then
'                            Qtde_entrada = Qtde_entrada - TBGravar!Quantidade
'                            TBGravar.Delete
'                        Else
'                            If Qtde_entrada > TBGravar!Quantidade Then
'                                Qtde_entrada = Qtde_entrada - TBGravar!Quantidade
'                                TBGravar!Quantidade = 0
'                            Else
'                                TBGravar!Quantidade = TBGravar!Quantidade - Qtde_entrada
'                                Qtde_entrada = 0
'                            End If
'                            TBGravar!Quantidade_PC = TBGravar!Quantidade
'                            TBGravar.Update
'                        End If
'                    End If
'                End If
'                TBGravar.Close
'                TBTempo.MoveNext
'            Loop
'        End If
'        TBTempo.Close
'        TBCFOP.MoveNext
'    Loop
'End If
'TBCFOP.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcRemoveListaResize(Formulario As Form)
On Error GoTo tratar_erro
Dim Controle As Control

For i = 0 To Formulario.Controls.Count - 1
    Set Controle = Formulario.Controls(i)
    If TypeOf Controle Is ListView Then Controle.Tag = "Ex_Columns, Ex_Font"
Next i

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcFunAbreBD_Configuracao(ServidorConf As String, BancoConf As String, UsuarioConf As String, SenhaConf As String)
On Error GoTo tratar_erro

Set Conexao_Configuracao = New adodb.Connection
With Conexao_Configuracao
    .Provider = "SQLOLEDB"
    .Properties("Data Source").Value = ServidorConf
    .Properties("Initial catalog").Value = BancoConf
    .Properties("User ID").Value = UsuarioConf
    .Properties("Password").Value = SenhaConf
    .Properties("Persist Security Info") = "False"
    .Open
    .Close
End With

Exit Sub
tratar_erro:
    If Err.Number = "-2147467259" Then
        Permitido = False
        Exit Sub
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunVerificaQtdeEmpenhoREOrdem(TextoFiltro As String, VerifQtdePC As Boolean) As Double
On Error GoTo tratar_erro
Dim CamposFiltro As String 'OK

FunVerificaQtdeEmpenhoREOrdem = 0
If VerifQtdePC = True Then CamposFiltro = "Qtde_empenhar_PC" Else CamposFiltro = "Qtde_empenhar"
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(" & CamposFiltro & ") as Qtde_requisitar from Qtde_empenhada_produto_detalhado where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    FunVerificaQtdeEmpenhoREOrdem = IIf(IsNull(TBAbrir!Qtde_requisitar), 0, TBAbrir!Qtde_requisitar)
End If
TBAbrir.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Function FunVerificaQtdeEmpenhoREPI(TextoFiltro As String) As Double
On Error GoTo tratar_erro

FunVerificaQtdeEmpenhoREPI = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(Qtde_requisitar) as Qtde_requisitar from Qtde_empenhada_produto_venda_detalhado where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    FunVerificaQtdeEmpenhoREPI = IIf(IsNull(TBAbrir!Qtde_requisitar), 0, TBAbrir!Qtde_requisitar)
End If
TBAbrir.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Function FunVerifQtdeOSProcControlado(OS As Long) As Boolean
On Error GoTo tratar_erro

FunVerifQtdeOSProcControlado = True
Set TBProducao = CreateObject("adodb.recordset")
TBProducao.Open "Select OS.IDProducao, OS.Fase, OS.TotalProd, OS.Quantidade, P.Ordem, P.QUANT from ordemservico OS INNER JOIN Producao P ON OS.Ordem = P.Ordem where OS.IDproducao = " & OS, Conexao, adOpenKeyset, adLockOptimistic
If TBProducao.EOF = False Then
    Produzidas = IIf(IsNull(TBProducao!Totalprod), 0, TBProducao!Totalprod)
    ProcVerifQtdePCSOSAnt TBProducao!Ordem, TBProducao!Fase, TBProducao!IDProducao, TBProducao!Quantidade
    
    'Verifica se a quantidade produzida é maior que a quantidade da OS anterior
    If Processo_controlado = True Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select IDproducao from ordemservico where Ordem = " & TBProducao!Ordem & " and Retrabalho = 'False' and Fase < '" & TBProducao!Fase & "' order by Fase desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If TBAbrir!IDProducao <> TBProducao!IDProducao And (Produzidas + TOK + TNC) > qtdeliberada Then
                MsgBox ("OS com PROCESSO CONTROLADO, a quantidade total apontada é maior que a quantidade apontada na OS anterior." & vbCrLf & vbCrLf & "OS anterior: " & qtdeliberada & vbCrLf & "Produzidas: " & Produzidas & vbCrLf & "+ Conforme: " & TOK & vbCrLf & "+ NC: " & TNC & vbCrLf & "= Total: " & Produzidas + TOK + TNC), vbExclamation
                FunVerifQtdeOSProcControlado = False
                TBAbrir.Close
                TBProducao.Close
                Exit Function
            End If
        Else
            If OSControlada = True Then
                If (Produzidas + TOK + TNC) > qtdeliberada Then
                    MsgBox ("ORDEM CONTROLADA, a quantidade total apontada é maior que a quantidade do lote." & vbCrLf & vbCrLf & "Lote: " & qtdeliberada & vbCrLf & "Produzidas: " & Produzidas & vbCrLf & "+ Conforme: " & TOK & vbCrLf & "+ NC: " & TNC & vbCrLf & "= Total: " & Produzidas + TOK + TNC), vbExclamation
                    FunVerifQtdeOSProcControlado = False
                    TBProducao.Close
                    Exit Function
                End If
            End If
        End If
    End If
    
    If OSControlada = True Then
        'Se for encerrar a ordem com quantidade menor que o programado
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select IDProducao from ordemservico where Ordem = " & TBProducao!Ordem & " and Retrabalho = 'False' and Fase < '" & TBProducao!Fase & "' and Pronto = 'NÃO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then qtdeliberada = TBProducao!Quant
        
        If Processo_controlado = False Then
            If (Produzidas + TOK + TNC) > qtdeliberada Then
                MsgBox ("ORDEM CONTROLADA, a quantidade total apontada é maior que a quantidade do lote." & vbCrLf & vbCrLf & "Lote: " & qtdeliberada & vbCrLf & "Produzidas: " & Produzidas & vbCrLf & "+ Conforme: " & TOK & vbCrLf & "+ NC: " & TNC & vbCrLf & "= Total: " & Produzidas + TOK + TNC), vbExclamation
                FunVerifQtdeOSProcControlado = False
                TBProducao.Close
                Exit Function
            End If
        End If
        
        If Evento = 3 Then
            If (Produzidas + TOK + TNC) < qtdeliberada Then
                MsgBox ("ORDEM CONTROLADA, a quantidade total apontada é menor que a quantidade do lote." & vbCrLf & vbCrLf & "Lote: " & qtdeliberada & vbCrLf & "Produzidas: " & Produzidas & vbCrLf & "+ Conforme: " & TOK & vbCrLf & "+ NC: " & TNC & vbCrLf & "= Total: " & Produzidas + TOK + TNC), vbExclamation
                TBProducao.Close
                FunVerifQtdeOSProcControlado = False
                Exit Function
            End If
        Else
            'Se for encerrar o evento com quantidade produzida igual ao lote não permite
            Permitido = False
            If Processo_controlado = True Then
                'Verifica se a OS apontada não é a primeira
                Set TBFiltro = CreateObject("adodb.recordset")
                TBFiltro.Open "Select IDProducao from ordemservico where Ordem = " & TBProducao!Ordem & " and Retrabalho = 'False' order by IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                If TBFiltro.EOF = False Then
                    If TBFiltro!IDProducao <> TBProducao!IDProducao Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select OS.* from " & NomeTabelaAp & " AP INNER JOIN ordemservico OS ON OS.IdProducao = AP.OS where AP.Ordem = " & TBProducao!Ordem & " and OS.Retrabalho = 'False' and AP.CodigoDesc = 3 and (AP.OS = " & TBProducao!IDProducao - 1 & " or AP.Fase = '" & TBProducao!Fase - 10 & "')", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Permitido = True
                        End If
                        TBAbrir.Close
                    Else
                        Permitido = True
                    End If
                End If
                TBFiltro.Close
            Else
                Permitido = True
            End If
            If Permitido = True Then
                If (Produzidas + TOK + TNC) = qtdeliberada Then
                    MsgBox ("ORDEM CONTROLADA, a quantidade total apontada é igual ao lote, só é permitido encerrar a ordem com o evento ***FIM DE PRODUÇÃO ***." & vbCrLf & vbCrLf & "Lote: " & qtdeliberada & vbCrLf & "Produzidas: " & Produzidas & vbCrLf & "+ Conforme: " & TOK & vbCrLf & "+ NC: " & TNC & vbCrLf & "= Total: " & Produzidas + TOK + TNC), vbExclamation
                    TBProducao.Close
                    FunVerifQtdeOSProcControlado = False
                    Exit Function
                End If
            End If
        End If
    End If
End If
TBProducao.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcVerifQtdePCSOSAnt(Ordem As Long, Fase As String, IDProducao As Long, Quantidade As Double)
On Error GoTo tratar_erro

If Processo_controlado = True Then
    'Se a OS tiver processo controlado não pode apontar a quantidade maior que a quantidade da OS anterior mesmo se ela não estiver concluída
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select QTOK,QTCD, Fase from ordemservico where Ordem = " & Ordem & " and Retrabalho = 'False' and Fase < '" & Fase & "' order by Fase desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qtdeliberada = IIf(IsNull(TBAbrir!QTOK), 0, TBAbrir!QTOK) + IIf(IsNull(TBAbrir!QTCD), 0, TBAbrir!QTCD) + FunVerifQtdeRetrabalhoFase(Ordem, TBAbrir!Fase)
    Else
        qtdeliberada = Quantidade
    End If
    TBAbrir.Close
Else
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(QTNC) as TTNC from ordemservico where Ordem = " & Ordem & " and Idproducao < " & IDProducao, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        QtdeRefugo = IIf(IsNull(TBAbrir!TTNC), 0, TBAbrir!TTNC)
    End If
    TBAbrir.Close
    qtdeliberada = Quantidade - QtdeRefugo
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunVerifNomeServidor() As String
On Error GoTo tratar_erro
Dim Numero As Long
Dim Numero1 As Long

Texto = ""
Numero = 0
Numero1 = Len(NomeServidor)
Hora = 0
If Numero1 <> 1 Then
    Do While Numero1 <> 0
        If Texto = "\" Then GoTo Pula
        Texto = Left(NomeServidor, (Numero + 1))
        Texto = Right(Texto, Len(Texto) - Numero)
        Numero = Numero + 1
        Numero1 = Numero1 - 1
    Loop
End If
Pula:
    FunVerifNomeServidor = Left(NomeServidor, Numero - 1)

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Function FunHoraServidor(ByVal pNomeServidor As String) As Variant
On Error GoTo tratar_erro
Dim t As TIME_OF_DAY
Dim tPtr As Long
Dim Resultado As Long
Dim szServer As String
Dim dataServidor As Date

If Left(pNomeServidor, 2) = "\\" Then szServer = StrConv(pNomeServidor, vbUnicode) Else szServer = StrConv("\\" & pNomeServidor, vbUnicode)
Resultado = NetRemoteTOD(szServer, tPtr)

If Resultado = 0 Then
    Call CopyMemory(t, ByVal tPtr, Len(t))
    dataServidor = DateSerial(70, 1, 1) + (t.t_elapsedt / 60 / 60 / 24)
    dataServidor = dataServidor - (t.t_timezone / 60 / 24)
    NetApiBufferFree (tPtr)
    FunHoraServidor = dataServidor
Else
    FunHoraServidor = Now
End If

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function FunProgBarKFTP(prog_bar As PictureBox, prog_hst As PictureBox, pc As Byte)
On Error GoTo tratar_erro

prog_bar.Width = Int((pc / 100) * prog_hst.Width)
prog_bar.Visible = (pc <> 0)
DoEvents
prog_bar.Refresh

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function FunConectaKFTP(kftp As kftp) As Boolean
On Error GoTo tratar_erro

FunConectaKFTP = True
If Not kftp.Connect("ftp.caprind.com.br", "caprind1", "cap0902loc", Val("21")) Then
    If kftp.LastError = "Please disconnect" Then kftp.Disconnect
    MsgBox ("Não foi possível conectar com o servidor de atualização, a atualização será encerrada."), vbExclamation
    FunConectaKFTP = False
Else
    If Not kftp.ChangeWorkingDir("public_html/phocadownload/userupload") Then
        MsgBox ("Não foi localizado a pasta no servidor de atualização, a atualização será encerrada."), vbExclamation
        FunConectaKFTP = False
    End If
End If

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function FunDownloadKFTP(kftp As kftp, NomeArquivoBaixar As String, CaminhoSalvar As String) As Boolean
On Error GoTo tratar_erro

FunDownloadKFTP = True
kftp.ChangeTransfertMode (Binary)
If Not kftp.DownloadFile(NomeArquivoBaixar, CaminhoSalvar, 0, True) Then
    MsgBox ("Não foi possível baixar a atualização, a atualização será encerrada."), vbExclamation
    FunDownloadKFTP = False
End If

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcImprimirDireto(FormulaRel As String, FormulaRelSubReport As String)
On Error GoTo tratar_erro

ProcVerifRelPersonalizado
            
If PermitidoRel = False Then LocalrelNovo = Localrel Else LocalrelNovo = LocalRelPersonalizado
Set Report = crAPP.OpenReport(LocalrelNovo & "\" & Nomerel)
'Login SQL
Contador = Report.Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop
ProcVerifSubReport FormulaRelSubReport

Report.FormulaSyntax = crCrystalSyntaxFormula 'Configura a sintaxe da formula
Report.RecordSelectionFormula = FormulaRel 'Formula de seleção do relatório
Report.PrintOut False 'Configura a seleção de impressora com false, enviando para impressora padrão
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    If Err.Number = "-2147206461" Then
        MsgBox ("Não foi encontrado o relatório " & Nomerel & " na pasta " & LocalrelNovo), vbExclamation
        GoTo 2
    End If
    If Err.Number = "-2147483638" Then
        MsgBox ("Não foi possível visualizar o relatório, favor reiniciar o sistema."), vbExclamation
        GoTo 2
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcLogonBDSQL()
On Error GoTo tratar_erro

Set CPProperty = DBTable.ConnectionProperties("Data Source")
CPProperty.Value = NomeServidor
Set CPProperty = DBTable.ConnectionProperties("User ID")
CPProperty.Value = IIf(Usuario_banco = "", "Procam", Usuario_banco)
Set CPProperty = DBTable.ConnectionProperties("Password")
CPProperty.Value = IIf(Senha_banco = "", "PRO0902loc$?", Senha_banco)
Set CPProperty = DBTable.ConnectionProperties("Initial Catalog")
CPProperty.Value = Nome_banco
DBTable.Location = "authors2"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcVerifSubReport(FormulaRelSubReport As String)
On Error GoTo tratar_erro

Contador2 = 0
NomeSubReport = ""
NomeSubReport1 = ""
NomeSubReport2 = ""
NomeSubReport3 = ""
NomeSubReport4 = ""
NomeSubReport5 = ""
NomeSubReport6 = ""
NomeSubReport7 = ""
NomeSubReport8 = ""
NomeSubReport9 = ""
NomeSubReport10 = ""

Set TBSubreport = CreateObject("adodb.recordset")
TBSubreport.Open "Select * from Qualidade_revisao_relatorios_subreports where Nome_relatorio = '" & Nomerel & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBSubreport.EOF = False Then
    Do While TBSubreport.EOF = False
        Select Case Contador2
            Case 0: NomeSubReport = TBSubreport!SubReport
            Case 1: NomeSubReport1 = TBSubreport!SubReport
            Case 2: NomeSubReport2 = TBSubreport!SubReport
            Case 3: NomeSubReport3 = TBSubreport!SubReport
            Case 4: NomeSubReport4 = TBSubreport!SubReport
            Case 5: NomeSubReport5 = TBSubreport!SubReport
            Case 6: NomeSubReport6 = TBSubreport!SubReport
            Case 7: NomeSubReport7 = TBSubreport!SubReport
            Case 8: NomeSubReport8 = TBSubreport!SubReport
            Case 9: NomeSubReport9 = TBSubreport!SubReport
            Case 10: NomeSubReport10 = TBSubreport!SubReport
        End Select
        Contador2 = Contador2 + 1
        TBSubreport.MoveNext
    Loop
End If
TBSubreport.Close
If Contador2 > 0 Then ProcSubReport FormulaRelSubReport

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSubReport(FormulaRelSubReport As String)
On Error GoTo tratar_erro
Dim SubReportRel As String 'OK

Do While Contador2 > 0
    Select Case Contador2
        Case "10": SubReportRel = NomeSubReport9
        Case "9": SubReportRel = NomeSubReport8
        Case "8": SubReportRel = NomeSubReport7
        Case "7": SubReportRel = NomeSubReport6
        Case "6": SubReportRel = NomeSubReport5
        Case "5": SubReportRel = NomeSubReport4
        Case "4": SubReportRel = NomeSubReport3
        Case "3": SubReportRel = NomeSubReport2
        Case "2": SubReportRel = NomeSubReport1
        Case "1": SubReportRel = NomeSubReport
    End Select
    
    Contador = Report.OpenSubreport(SubReportRel).Database.Tables.Count
    Do While Contador > 0
        Set DBTable = Report.OpenSubreport(SubReportRel).Database.Tables(Contador)
        ProcLogonBDSQL
        Contador = Contador - 1
        
        'Coloca a formula no subreport
        If FormulaRelSubReport <> "" And SubReportRel <> "RevisaoRelatorio.rpt" And SubReportRel <> "Responsavel_relatorio" Then
            Report.OpenSubreport(SubReportRel).FormulaSyntax = crCrystalSyntaxFormula
            Report.OpenSubreport(SubReportRel).RecordSelectionFormula = FormulaRelSubReport
        End If
        
        'Coloca a formula no subreport de responsavel
'        If SubReportRel = "Responsavel_relatorio" Then
'            'verifica se esta marcado para imprimir o relatorio com o responsavel
'            Set TBMaterial = CreateObject("adodb.recordset")
'            TBMaterial.Open "Select Responsavel_rel from Qualidade_revisao_relatorios where Nome_relatorio = '" & Nomerel & "' and Responsavel_rel = 1", Conexao, adOpenKeyset, adLockOptimistic
'            If TBMaterial.EOF = False Then
'                Report.OpenSubreport(SubReportRel).FormulaSyntax = crCrystalSyntaxFormula
'                Report.OpenSubreport(SubReportRel).RecordSelectionFormula = "{Usuarios.IDusuario} = " & pubIDUsuario
'            End If
'            TBMaterial.Close
'        End If
    Loop
1:
    Contador2 = Contador2 - 1
Loop

Exit Sub
tratar_erro:
    If Err.Number = "-2147190528" Then GoTo 1
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Function FunTamanhoTextoZeroEsq(Texto As Variant, Tamanho As Integer) As String
On Error GoTo tratar_erro
Dim QuantZeroEsq As Double 'OK
Dim Texto1 As String 'OK

Texto1 = ""
QuantZeroEsq = Tamanho - Len(Texto)
If QuantZeroEsq > 0 Then
    Do While QuantZeroEsq > 0
        If Texto1 = "" Then Texto1 = "0" Else Texto1 = Texto1 & "0"
        QuantZeroEsq = QuantZeroEsq - 1
    Loop
    FunTamanhoTextoZeroEsq = Texto1 & Texto
Else
    FunTamanhoTextoZeroEsq = Texto
End If

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcVerifQtdeLicencas()
On Error GoTo tratar_erro

Qtlicencas_gerprod = 0
If TemInternet = True And ErroDriverMYSQL = False Then
    Set TBFiltro = CreateObject("adodb.recordset")
    TBFiltro.Open "Select CNPJ from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBFiltro.EOF = False Then
        FunAbreBDSite
        If ConexaoMySql.State = 1 Then
            Set TBMySQL = New adodb.Recordset
            TBMySQL.Open "Select Licencas_gerprod From Clientes Where CNPJ = '" & TBFiltro!CNPJ & "'", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
            With TBMySQL
                If .EOF = False Then
                    'Verifica número de licenças
                    If IsNull(.Fields!Licencas_gerprod) = False And .Fields!Licencas_gerprod <> "" Then Qtlicencas_gerprod = .Fields!Licencas_gerprod
                End If
            End With
        End If
    End If
    TBFiltro.Close
Else
    Set TBFiltro = CreateObject("adodb.recordset")
    TBFiltro.Open "Select Licencas_gerprod from Empresa", Conexao, adOpenKeyset, adLockOptimistic
    If TBFiltro.EOF = False Then
        If IsNull(TBFiltro!Licencas_gerprod) = False And TBFiltro!Licencas_gerprod <> "" Then Qtlicencas_gerprod = TBFiltro!Licencas_gerprod
    End If
    TBFiltro.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunProcCalculaQtdePC(Codinterno As String, quantidadeUN As Double, CalculaPC As Boolean, Unconversao As String) As Double
On Error GoTo tratar_erro

FunProcCalculaQtdePC = 0
If Codinterno = "" Then Exit Function
Set TBUN = CreateObject("adodb.recordset")
TBUN.Open "Select Un_Kg, PBruto, Unidade from projproduto where Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUN.EOF = False Then
    If TBUN!Unidade = "PÇ" Or TBUN!Unidade = "PC" Or TBUN!Unidade = "UN" Or TBUN!Unidade = "CJ" Then
        'Calcula quantidade se a unidade for diferente
        If TBUN!Unidade <> Unconversao And Unconversao <> "" Then
            If FunVerifUNConversao(TBUN!Unidade, Unconversao) = True Then
                FunProcCalculaQtdePC = FunConverteUN(TBUN!Unidade, Unconversao, quantidadeUN, Codinterno)
            Else
                FunProcCalculaQtdePC = quantidadeUN / FunVerificaConversaoUnidade(TBUN!Unidade, Unconversao)
            End If
        Else
            FunProcCalculaQtdePC = quantidadeUN
        End If
    Else
        If Unconversao = "PÇ" Or Unconversao = "PC" Or Unconversao = "UN" Or Unconversao = "CJ" Then
            FunProcCalculaQtdePC = quantidadeUN
        Else
            If CalculaPC = True Then
                If TBUN!Unidade = "KG" And IsNull(TBUN!PBruto) = False And TBUN!PBruto <> 0 And (TBUN!Un_Kg = "Mt²" Or TBUN!Un_Kg = "Mt/L") Then FunProcCalculaQtdePC = Format(quantidadeUN / TBUN!PBruto, "###,##0.0000")
            Else
                FunProcCalculaQtdePC = Format(quantidadeUN * TBUN!PBruto, "###,##0.0000")
            End If
        End If
    End If
End If
TBUN.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

Function FunVerifUNConversao(Un_est As String, Un_com As String) As Boolean
On Error GoTo tratar_erro

FunVerifUNConversao = False
If Un_est <> Un_com And (Un_est = "KG" Or Un_est = "MT" Or Un_est = "MM" Or Un_est = "BR" Or Un_est = "PC" Or Un_est = "PÇ" Or Un_est = "CH") And (Un_com = "KG" Or Un_com = "MT" Or Un_com = "MM" Or Un_com = "BR" Or Un_com = "PC" Or Un_com = "PÇ" Or Un_com = "CH") Then FunVerifUNConversao = True

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

Function FunConverteUN(Un_est As String, Un_com As String, quantidadeUN1 As Double, DesenhoUN As String) As Double
On Error GoTo tratar_erro
Dim quantidadeUN2 As Double

Set TBUN1 = CreateObject("adodb.recordset")
TBUN1.Open "Select peso_metro, PBruto from projproduto where desenho = '" & DesenhoUN & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUN1.EOF = False Then
    If Un_est = "KG" Then
        If Un_com = "PÇ" Or Un_com = "PC" Or Un_com = "BR" Or Un_com = "CH" Then quantidadeUN2 = IIf(IsNull(TBUN1!PBruto), 0, TBUN1!PBruto) Else quantidadeUN2 = IIf(IsNull(TBUN1!peso_metro), 0, TBUN1!peso_metro)
        If Un_com = "MT" Then
            FunConverteUN = Format(quantidadeUN2 * quantidadeUN1, "###,##0.0000000000")
        ElseIf Un_com = "MM" Then
                FunConverteUN = Format((quantidadeUN2 / 1000) * quantidadeUN1, "###,##0.0000000000")
            ElseIf Un_com = "PÇ" Or Un_com = "PC" Or Un_com = "BR" Or Un_com = "CH" Then
                FunConverteUN = Format(quantidadeUN2 * quantidadeUN1, "###,##0.0000000000")
        End If
    Else
        If Un_est = "PÇ" Or Un_est = "PC" Or Un_est = "BR" Or Un_est = "CH" Then quantidadeUN2 = IIf(IsNull(TBUN1!PBruto), 0, TBUN1!PBruto) Else quantidadeUN2 = IIf(IsNull(TBUN1!peso_metro), 0, TBUN1!peso_metro)
        If Un_est = "MT" Then
            FunConverteUN = Format(quantidadeUN1 / quantidadeUN2, "###,##0.0000000000")
        ElseIf Un_est = "MM" Then
                FunConverteUN = Format((quantidadeUN1 * 0.001) / quantidadeUN2, "###,##0.0000000000")
            ElseIf Un_est = "PÇ" Or Un_est = "PC" Or Un_est = "BR" Or Un_est = "CH" Then
                FunConverteUN = Format(quantidadeUN1 / quantidadeUN2, "###,##0.0000000000")
        End If
    End If
End If
TBUN1.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

Public Function FunVerificaConversaoUnidade(Unidade_de As String, Unidade_para As String) As Double
On Error GoTo tratar_erro

FunVerificaConversaoUnidade = 1
If Unidade_de <> Unidade_para Then
    Set TBUN2 = CreateObject("adodb.recordset")
    TBUN2.Open "Select * from Tabela_conversao_unidade where Unidade_de = '" & Unidade_de & "' and Unidade_para = '" & Unidade_para & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUN2.EOF = False Then
        FunVerificaConversaoUnidade = IIf(IsNull(TBUN2!Qtde_para), 1, TBUN2!Qtde_para)
    End If
    TBUN2.Close
End If

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

Public Function FunVerifNCPorDescricao(OS As Long) As Boolean
On Error GoTo tratar_erro

FunVerifNCPorDescricao = False
Set TBUN = CreateObject("adodb.recordset")
TBUN.Open "Select E.Codigo from (OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem) INNER JOIN Empresa E ON E.Codigo = P.ID_empresa where OS.Idproducao = " & OS & " and E.Apontar_NC_descricao = 1", Conexao, adOpenKeyset, adLockReadOnly
If TBUN.EOF = False Then
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from cadmaquinas where maquina = '" & frmProducao.txtMaquina.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
If TBAbrir!Insp_final = True Then
    FunVerifNCPorDescricao = True
End If
TBAbrir.Close
End If
End If
TBUN.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

Public Function FunVerifNCRejeitado(OS As Long) As Boolean
On Error GoTo tratar_erro

FunVerifNCRejeitado = False
Set TBUN = CreateObject("adodb.recordset")
TBUN.Open "Select E.Codigo from (OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem) INNER JOIN Empresa E ON E.Codigo = P.ID_empresa where OS.Idproducao = " & OS & " and E.NC_parecer_rejeitado = 1", Conexao, adOpenKeyset, adLockReadOnly
If TBUN.EOF = False Then
    FunVerifNCRejeitado = True
End If
TBUN.Close

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function
