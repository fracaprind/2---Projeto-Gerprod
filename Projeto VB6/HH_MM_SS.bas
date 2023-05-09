Attribute VB_Name = "Mdl_HH_MM_SS"
Public Valor            As Double
Public Valor1           As Double
Public Valor2           As Double
Public Valor3           As Double

Public TotalH As String
Public TotalM As String
Public TotalS As String

'=================================================
'    VARIAVEIS DE CALCULO DE TEMPOS DE PROCESSO  =
'=================================================
Public DataResultado As Date
Public HoraResultado As Date
Public TempoPreparacao As String
Public TempoExecucao As String
Public Preparacao As Variant
Public Execucao As Variant

'===========================
'=   VARIAVEIS DE DATAS    =
'===========================
Public Data As Date
Public DataConclusaoOS As Date
Public DataConclusaoOrdem As Date
'===========================
'=   VARIAVEIS DE DATAS    =
'===========================
Public PDTETOT As Date
Public PRTETOT As Date
Public TA As Date
Public TB As Date
Public HA As Date
Public HJ As Date
Public TEMPODISP As Date 'Totalizacao de Tempo total disponivel
Public Dataini As Date
Public OSControlada As Boolean 'Controle de Ordem de serviço controlada sim/não
Public Processo_controlado As Boolean

Public ULTICOD As Integer 'Ultimo código do evento da lista
Public ULTIDESC As String 'Ultimo evento da lista
Public ULTIOPERADOR 'Ultimo operador a apontar evento da O.s
Public TEP As Date 'Tempo de execucao previsto
Public QTLOTE As Long 'Quantidade do lote
Public TPP As Date 'Tempo previsto por peça

'Public DataVariavel As Date
'Public TotalDias As Date
Public ValorDia As Date

Public Dia As Integer
Public Dias As Integer
Public Hora As Double
Public Minuto As Double
Public Segundo As Double

Public Segprev As Double 'Segundos utilizados
Public SegUtil As Double 'Segundos previstos

Public D, H, M, S As Double ' Controles da função FunElapsedTime
Public Horas As Integer ' Controles da função FunElapsedTime
Public Minutos As Integer ' Controles da função FunElapsedTime
Public Segundos As Double ' Controles da função FunElapsedTime
Public TotalDia As Date ' Controles da função FunElapsedTime
Public TotalHora As Date ' Controles da função FunElapsedTime
Public TotalMinuto As Date ' Controles da função FunElapsedTime
Public TotalSegundo As Date ' Controles da função FunElapsedTime
Public Horatotal As String 'Total em horas minutos e segundos retornados pela função FunElapsedTime ( 123:12:33 )

Public TempoInicio As Date 'Hora de inicio do evento
Public TempoFinal As Date 'Hora final do evento
Public TempoTotal As Date 'Total de tempo do evento

Public TempoTotalProd As Date 'Somatório de tempo total produzindo
Public TempoTotalPrep As Date 'Somatório de tempo total preparando a máquina

Public TotalSegundos As Double 'Tempo total em segundos
Public TotalDisponivel As Date
Public TempoUltimo As Date

Public Function FunFormataTempo(TotalSeg As Double)
On Error GoTo tratar_erro

totalHoras = 0
TotalMin = 0
Do While TotalSeg >= 60
    TotalMin = TotalMin + 1
    TotalSeg = TotalSeg - 60
Loop
If TotalSeg = 60 Then
    TotalSeg = 0
    TotalMin = TotalMin + 1
End If
Do While TotalMin >= 60
    totalHoras = totalHoras + 1
    TotalMin = TotalMin - 60
Loop

TotalSeg = Format(TotalSeg, "###,##0.0000000000")
If TotalSeg < 10 Then
    FunFormataTempo = IIf(Len(totalHoras) < 2, "0" & totalHoras, totalHoras) & ":" & IIf(Len(TotalMin) = 2, TotalMin, "0" & TotalMin) & ":" & IIf(Len(TotalSeg) = 2, TotalSeg, "0" & TotalSeg)
Else
    FunFormataTempo = IIf(Len(totalHoras) < 2, "0" & totalHoras, totalHoras) & ":" & IIf(Len(TotalMin) = 2, TotalMin, "0" & TotalMin) & ":" & TotalSeg
End If

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Function FunElapsedTime(Interval)
On Error GoTo tratar_erro

D = Int(CSng(Interval))
H = Format(Int(CSng(Interval * 24)), "###,###,##0")
M = Format(Int(CSng(Interval * 24 * 60)), "###,###,##0")
S = Format(Int(CSng(Interval * 24 * 3600)), "###,###,##0")

'Debug.Print "Dia(s) = " & D
'Debug.Print "hora(s) = " & H
'Debug.Print "minuto(s) = " & M
'Debug.Print "Segundo(s) = " & S

Horas = H
Minutos = M Mod 60
Segundos = S Mod 60

Hr = Horas
Mn = Minutos
Sg = Segundos

Horatotal = IIf(Len(Hr) = 1, "0" & Hr, Hr) & ":" & IIf(Len(Mn) = 1, "0" & Mn, Mn) & ":" & IIf(Len(Sg) = 1, "0" & Sg, Sg)

'Debug.Print Horatotal

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Sub ProcFormataHora(HoraFormato As String)
On Error GoTo tratar_erro
Dim Hora As Long 'OK
Dim DataMinuto As Date 'OK
Dim DataMinuto1 As Date 'OK

DataResultado = 0
DecimoSegundos = 0
Texto = ""
Numero = 0
Numero1 = Len(HoraFormato)
Hora = 0
If Numero1 <> 1 Then
    Do While Numero1 <> 0
        If Texto = ":" Then GoTo Pula
        Texto = Left(HoraFormato, (Numero + 1))
        Texto = Right(Texto, Len(Texto) - Numero)
        Numero = Numero + 1
        Numero1 = Numero1 - 1
    Loop
Pula:
    Hora = Left(HoraFormato, (Numero - 1))
    Texto1 = Hora
    Numero2 = Len(Texto1)
    If Hora >= 24 Then
        Do While Hora >= 24
            Hora = Hora - 24
            DataResultado = DataResultado + #11:59:59 PM# + #12:00:01 AM#
        Loop
    End If
        
    'Verifica qtde. de horas
    Texto = ""
    Numero = 0
    Numero1 = Len(HoraFormato)
    Do While Numero1 <> 0
        If Texto = ":" Then GoTo Pula1
        Texto = Left(HoraFormato, (Numero + 1))
        Texto = Right(Texto, Len(Texto) - Numero)
        Numero = Numero + 1
        Numero1 = Numero1 - 1
    Loop
Pula1:
    MinutoSeg = Right(HoraFormato, Len(HoraFormato) - Numero)
    If Len(MinutoSeg) = 5 Then
        DataMinuto = FunFormataTempo(Right(MinutoSeg, 2))
        DataMinuto1 = "00:" & Left(MinutoSeg, 2) & ":00"
        MinutoSeg = Right(DataMinuto + DataMinuto1, 5)
    End If
    If Hora < 10 Then Texto = "0" & Hora & ":" & MinutoSeg Else Texto = Hora & ":" & MinutoSeg
    DecimoSegundos = IIf(Len(Texto) > 8, Right(Texto, Len(Texto) - 8), 0)
    If DataResultado <> "00:00:00" Then
        If Hora < 10 Then DataResultado = DataResultado & " " & "0" & Hora & ":" & Left(MinutoSeg, 5) Else DataResultado = DataResultado & " " & Hora & ":" & Left(MinutoSeg, 5)
    Else
        DataResultado = Left(Texto, 8)
    End If
End If
    FunElapsedTime (DataResultado)
    'Debug.Print Horatotal

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Function FunCalculaSegPC(TExec As String, PcHora As Double)
On Error GoTo tratar_erro

If IsDate(TExec) = True Then
    Dataini = TExec
    FunElapsedTime (Dataini)
Else
    ProcFormataHora (TExec)
End If
Valor1 = S
Valor2 = PcHora
If Valor1 And Valor2 <> 0 Then FunCalculaSegPC = Valor1 / Valor2 Else FunCalculaSegPC = 0

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Function FunSóNumeros(X As String) As String
On Error GoTo tratar_erro
Dim temp As String
Dim j As Integer

temp = ""
For j = 1 To Len(X)
    If Mid(X, j, 1) = "0" Or _
        Mid(X, j, 1) = "1" Or _
        Mid(X, j, 1) = "2" Or _
        Mid(X, j, 1) = "3" Or _
        Mid(X, j, 1) = "4" Or _
        Mid(X, j, 1) = "5" Or _
        Mid(X, j, 1) = "6" Or _
        Mid(X, j, 1) = "7" Or _
        Mid(X, j, 1) = "8" Or _
        Mid(X, j, 1) = "9" Then
        temp = temp + Mid(X, j, 1)
    End If
Next
numeros = temp
FunSóNumeros = temp

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function
