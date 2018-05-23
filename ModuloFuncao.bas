Attribute VB_Name = "ModuloFuncao"
Global VGStrDataErro As String
Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long

'declarações para saber se está conectado a internet
Private Const CONNECT_LAN As Long = &H2
Private Const CONNECT_MODEM As Long = &H1
Private Const CONNECT_PROXY As Long = &H4
Private Const CONNECT_OFFLINE As Long = &H20
Private Const CONNECT_CONFIGURED As Long = &H40

'declarações para fechar um aplicativo
Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
Declare Function PostMessage Lib "user32" Alias _
"PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10

Function VerificaNulo(vNulo)
    If IsNull(vNulo) = True Or vNulo = "" Then
        VerificaNulo = ""
    Else
        VerificaNulo = vNulo
    End If
End Function

Function FormataDataUS(vData)
 
    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    If vData = "" Or IsNull(vData) = True Or IsDate(vData) = False Then
        FormataDataUS = Null
    Else
        Dia = DatePart("D", vData)
        Mes = DatePart("M", vData)
        Ano = DatePart("YYYY", vData)
         
        If Dia < 10 Then
         Dia = "0" & Dia
        End If
         
        If Mes < 10 Then
         Mes = "0" & Mes
        End If
         
        FormataDataUS = Ano & "/" & Mes & "/" & Dia
    End If
End Function

Function FormataData(vData)
 
    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    If vData = "" Or IsNull(vData) = True Or IsDate(vData) = False Then
        FormataData = ""
    Else
        Dia = DatePart("D", vData)
        Mes = DatePart("M", vData)
        Ano = DatePart("YYYY", vData)
         
        If Dia < 10 Then
         Dia = "0" & Dia
        End If
         
        If Mes < 10 Then
         Mes = "0" & Mes
        End If
        
        FormataData = Dia & "/" & Mes & "/" & Ano
    End If
End Function

Function FormataDataEspecial(vData)
 
    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    If vData = "" Or IsNull(vData) = True Then
        FormataDataEspecial = ""
    Else
        If IsDate(vData) = True Then
            Dia = DatePart("D", vData)
            Mes = DatePart("M", vData)
            Ano = DatePart("YYYY", vData)
        
            If Dia < 10 Then
             Dia = "0" & Dia
            End If
             
            If Mes < 10 Then
             Mes = "0" & Mes
            End If
        
        Else
            Mes = Trim(Mid(Trim(vData), 1, 3))
            vData = Trim(Mid(Trim(vData), 4))
            
            Dia = Trim(Mid(Trim(vData), 1, 2))
            vData = Trim(Mid(Trim(vData), 3))
            
            Ano = Trim(Mid(Trim(vData), 1, 4))
        
            If Dia < 10 Then
             Dia = "0" & Dia
            End If
             
            If Mes = "Jan" Or Mes = "jan" Then
             Mes = "01"
            ElseIf Mes = "Fev" Or Mes = "fev" Or Mes = "Feb" Or Mes = "feb" Then
             Mes = "02"
            ElseIf Mes = "Mar" Or Mes = "mar" Then
             Mes = "03"
            ElseIf Mes = "Abr" Or Mes = "abr" Or Mes = "Apr" Or Mes = "apr" Then
             Mes = "04"
            ElseIf Mes = "Mai" Or Mes = "mai" Or Mes = "May" Or Mes = "may" Then
             Mes = "05"
            ElseIf Mes = "Jun" Or Mes = "jun" Then
             Mes = "06"
            ElseIf Mes = "Jul" Or Mes = "jul" Then
             Mes = "07"
            ElseIf Mes = "Ago" Or Mes = "ago" Or Mes = "Aug" Or Mes = "aug" Then
             Mes = "08"
            ElseIf Mes = "Set" Or Mes = "set" Or Mes = "Sep" Or Mes = "sep" Then
             Mes = "09"
            ElseIf Mes = "Out" Or Mes = "out" Or Mes = "Oct" Or Mes = "oct" Then
             Mes = "10"
            ElseIf Mes = "Nov" Or Mes = "nov" Then
             Mes = "11"
            ElseIf Mes = "Dez" Or Mes = "dez" Or Mes = "Dec" Or Mes = "dec" Then
             Mes = "12"
            End If
        
        End If
        
        FormataDataEspecial = Dia & "/" & Mes & "/" & Ano
    
    End If
    
End Function

Function FormataHora(vHora)
 
    Dim hora As String
    Dim min As String
    
    hora = DatePart("H", vHora)
    min = DatePart("N", vHora)
    
    If hora < 10 Then
     hora = "0" & hora
    End If
    
    If min < 10 Then
     min = "0" & min
    End If
     
    FormataHora = hora & ":" & min
  
End Function

Function FormataNum(vNum)
 
    Dim num As String
    
    If vNum <> "" And IsNull(vNum) = False Then
        If vNum < 10 Then
            num = 0 & vNum
        Else
            num = Val(vNum)
        End If
        
        FormataNum = num
    Else
        FormataNum = ""
    End If
  
End Function

Function FormataNumDec(vNum)
    
    If vNum = "" Then
        FormataNumDec = ""
    Else
        If InStr(vNum, ",") = 0 And InStr(vNum, ".") = 0 Then
            FormataNumDec = vNum & ".00"
        Else
            If InStr(vNum, ",") <> 0 Then
                FormataNumDec = Replace(vNum, ",", ".")
            ElseIf InStr(vNum, ".") <> 0 Then
                FormataNumDec = vNum
            End If
            If Len(Mid(FormataNumDec, InStr(FormataNumDec, ".") + 1)) = 1 Then
                FormataNumDec = FormataNumDec & "0"
            End If
        End If
        If InStr(FormataNumDec, "+") <> 0 Then
            FormataNumDec = Mid(FormataNumDec, 2)
        End If
    End If
    
End Function

Function FormataNumDecRec(vNum)
    
    If vNum = "" Then
        FormataNumDecRec = ""
    Else
        If InStr(vNum, ",") = 0 And InStr(vNum, ".") = 0 Then
            FormataNumDecRec = vNum & ".00"
        Else
            If InStr(vNum, ",") <> 0 Then
                FormataNumDecRec = Replace(vNum, ",", ".")
            ElseIf InStr(vNum, ".") <> 0 Then
                FormataNumDecRec = vNum
            End If
            If Len(Mid(FormataNumDecRec, InStr(FormataNumDecRec, ".") + 1)) = 1 Then
                FormataNumDecRec = FormataNumDecRec & "0"
            End If
        End If
        If InStr(FormataNumDecRec, "-") = 0 And InStr(FormataNumDecRec, "+") = 0 Then
            FormataNumDecRec = "+" & FormataNumDecRec
        End If
    End If
    
End Function

Function ArredondaNumDec(vNum)
    Dim IntTemp As Integer
    Dim DecTemp As Integer
    
    If vNum = "" Then
        ArredondaNumDec = ""
    Else
        If Len(Mid(vNum, InStr(vNum, ",") + 1)) = 1 Then
            ArredondaNumDec = vNum & "0"
            
        ElseIf Len(Mid(vNum, InStr(vNum, ",") + 1)) = 2 Then
            ArredondaNumDec = vNum
            
        ElseIf Len(Mid(vNum, InStr(vNum, ",") + 1)) >= 3 Then
            
            If Mid(vNum, InStr(vNum, ",") + 3) <= 5 Then
                ArredondaNumDec = Mid(vNum, 1, InStr(vNum, ",") + 2)
                
            ElseIf Mid(vNum, InStr(vNum, ",") + 3) > 5 Then
                IntTemp = Mid(vNum, 1, InStr(vNum, ",") - 1)
                DecTemp = Mid(vNum, InStr(vNum, ",") + 1, 2) + 1
                ArredondaNumDec = IntTemp & "," & DecTemp
            End If
        End If

    End If
    
End Function

Function FormataEixo(vEx)
    If vEx = "" Then
        FormataEixo = ""
    Else
        If InStr(vEx, "º") = 0 Then
            FormataEixo = vEx & "º"
        Else
            FormataEixo = vEx
        End If
    End If
End Function

Function FormataMoeda(pvalor)
    Dim valor As String
    Dim centavo As String
    Dim centavotemp As String
    Dim poscentavo As Integer
    Dim posreal As Integer
    Dim realtemp As String
    Dim real As String
    Dim lenreal As Integer
    Dim ponto As String
    Dim sinal As String
    
    If InStr(pvalor, "R$") = 0 Then
        pvalor = pvalor
    Else
        pvalor = Trim(Mid(pvalor, 3))
    End If
    
    If InStr(pvalor, "-") <> 0 Then
        sinal = Mid(pvalor, 1, 1)
        pvalor = Mid(pvalor, 2)
    End If
    
    valor = pvalor
    poscentavo = InStr(valor, ",")
    
    If poscentavo <> 0 Then
        centavotemp = Mid(valor, poscentavo)
        If Len(centavotemp) = 2 Then
            centavo = centavotemp & "0"
        Else
            centavo = centavotemp
        End If
        
        realtemp = Mid(valor, 1, poscentavo - 1)
    Else
        centavo = ",00"
        realtemp = valor
    End If
    
    posreal = Len(realtemp)
    lenreal = 3

    Do While posreal <> 0
        real = Mid(realtemp, posreal, 1) & real
        
        If Len(real) = lenreal Then
            If Len(realtemp) <> 3 Then
                If Mid(realtemp, Len(realtemp) - 3, 1) <> "." Then
                    real = "." & real
                End If
            Else
                real = "." & real
            End If
            lenreal = Len(real) + 3
        End If
        
        posreal = posreal - 1
    Loop
    
    ponto = Mid(real, 1, 1)
    If ponto = "." Then
        real = Mid(real, 2)
    End If
        
    If sinal = "-" Then
        FormataMoeda = "R$ " & sinal & real & centavo
    Else
        FormataMoeda = "R$ " & real & centavo
    End If
End Function

Function VerificaData(vData)
    Dim DiaTemp As String
    Dim MesTemp As String
    Dim AnoTemp As String

    If vData <> "__/__/____" Then
        vData = Replace(vData, "/", "")
        vData = Replace(vData, "-", "")
    Else
        VerificaData = "__/__/____"
        Exit Function
    End If
    
    If Len(vData) = 8 Then
        DiaTemp = Mid(vData, 1, 2)
        MesTemp = Mid(vData, 3, 2)
        AnoTemp = Mid(vData, 5)

        If DiaTemp > 31 Or MesTemp > 12 Then
            VPStrBox = MsgBox("Data inválida.", vbCritical, "Pró Vendas 2004 - Erro")
            VGStrDataErro = "sim"
        
        ElseIf IsDate(DiaTemp & "/" & MesTemp & "/" & AnoTemp) = False Then
            VPStrBox = MsgBox("Data inválida.", vbCritical, "Pró Vendas 2004 - Erro")
            VGStrDataErro = "sim"
            
        Else
            VerificaData = DiaTemp & "/" & MesTemp & "/" & AnoTemp
        End If
        
    ElseIf Len(vData) = 10 And InStr(vData, "/") And vData <> "__/__/____" Then
        DiaTemp = Mid(vData, 1, 2)
        MesTemp = Mid(vData, 4, 2)
        AnoTemp = Mid(vData, 7)

        If DiaTemp > 31 Or MesTemp > 12 Then
            VPStrBox = MsgBox("Data inválida.", vbCritical, "Pró Vendas 2004 - Erro")
            VGStrDataErro = "sim"
        
        ElseIf IsDate(DiaTemp & "/" & MesTemp & "/" & AnoTemp) = False Then
            VPStrBox = MsgBox("Data inválida.", vbCritical, "Pró Vendas 2004 - Erro")
            VGStrDataErro = "sim"
        
        Else
            VerificaData = DiaTemp & "/" & MesTemp & "/" & AnoTemp
        End If

    Else
        VPStrBox = MsgBox("Formato da data está incorreto.", vbCritical, "Pró Vendas 2004 - Erro")
        VGStrDataErro = "sim"
    End If
End Function

Function FormataDataCompleta(vData)
 
    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    If vData = "" Or IsNull(vData) = True Then
        FormataDataCompleta = ""
    Else
        Dia = DatePart("D", vData)
        Mes = DatePart("M", vData)
        Ano = DatePart("YYYY", vData)
    
        If Dia < 10 Then
         Dia = "0" & Dia
        End If
         
        If Mes = 1 Then
         Mes = "janeiro"
         
        ElseIf Mes = 2 Then
         Mes = "fevereiro"
         
        ElseIf Mes = 3 Then
         Mes = "março"
         
        ElseIf Mes = 4 Then
         Mes = "abril"
         
        ElseIf Mes = 5 Then
         Mes = "maio"
         
        ElseIf Mes = 6 Then
         Mes = "junho"
         
        ElseIf Mes = 7 Then
         Mes = "julho"
         
        ElseIf Mes = 8 Then
         Mes = "agosto"
         
        ElseIf Mes = 9 Then
         Mes = "setembro"
         
        ElseIf Mes = 10 Then
         Mes = "outubro"
         
        ElseIf Mes = 11 Then
         Mes = "novembro"
         
        ElseIf Mes = 12 Then
         Mes = "dezembro"
         
        End If
        
        FormataDataCompleta = Dia & " de " & Mes & " de " & Ano
    
    End If
    
End Function

Function FormataSemana(vData)
 
    Dim Sem As String
    
    If vData = "" Or IsNull(vData) = True Then
        FormataSemana = ""
    Else
        Sem = DatePart("W", vData)
        
        If Sem = 1 Then
         FormataSemana = "Domingo"
         
        ElseIf Sem = 2 Then
         FormataSemana = "Segunda"
         
        ElseIf Sem = 3 Then
         FormataSemana = "Terça"
         
        ElseIf Sem = 4 Then
         FormataSemana = "Quarta"
         
        ElseIf Sem = 5 Then
         FormataSemana = "Quinta"
         
        ElseIf Sem = 6 Then
         FormataSemana = "Sexta"
         
        ElseIf Sem = 7 Then
         FormataSemana = "Sábado"
         
        End If
    
    End If
    
End Function

Function PegarResolucao()
  Dim xTwips%, yTwips%, xPixels#, YPixels#
  xTwips = Screen.TwipsPerPixelX
  yTwips = Screen.TwipsPerPixelY
  YPixels = Screen.Height / yTwips
  xPixels = Screen.Width / xTwips
  VGStrResolucao = xPixels & "x" & YPixels
End Function

Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
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
IsWebConnected = WebTest
End Function


