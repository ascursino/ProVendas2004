VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000F&
   Caption         =   "Pró Vendas 2004 - Sistema Integrado de Vendas"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10545
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2760
      OleObjectBlob   =   "MDIPrincipal.frx":0CCA
      Top             =   1080
   End
   Begin VB.Menu MNUSobre 
      Caption         =   "&Sobre"
   End
   Begin VB.Menu MNUAtu 
      Caption         =   "&Atualização"
      Visible         =   0   'False
   End
   Begin VB.Menu MNUHelp 
      Caption         =   "&Help"
      Index           =   0
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrUpdate As String
Private m_lngDocSize As Long


Private Sub MDIForm_Load()
    Dim VLStrRegistro As String
    
    If FrmSplash.SHLocker1.SouRegistrado = True Then
        VLStrRegistro = "Registrado"
    Else
        VLStrRegistro = "Trial"
    End If
    
    Screen.MousePointer = vbNormal
    Me.Caption = "Pró Vendas 2004 - Sistema Integrado de Vendas - V." & App.Major & "." & App.Minor & "." & App.Revision & " (" & VLStrRegistro & ")"
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    If VLStrRegistro = "Registrado" Then
        MNUAtu.Enabled = True
    Else
        MNUAtu.Enabled = False
    End If
    
    'FrmIdentifica.Show
    ''MDIPrincipal.Caption = "Pró Vendas 2004 - Sistema Integrado de Vendas"
    
    FrmPrincipal.Show
End Sub

Private Sub MDIForm_Terminate()
    Unload Me
End Sub

Private Sub MNUAtu_Click()
    Screen.MousePointer = vbHourglass
    Unload FrmPrincipal

    FrmUpdate.Show
    
    Call FazUpdate
    
    Screen.MousePointer = vbNormal
End Sub
  
Private Sub MNUHelp_Click(Index As Integer)
    Dim i&
    i& = ShellExecute(0, "open", App.Path & "\ProVendas2004.chm", "", App.Path, SW_SHOW)
End Sub

Private Sub MNUSobre_Click()
    frmSistema.Show
End Sub

Public Function FazUpdate()
    'iniciando a verificação da conexão com a internet
    FrmUpdate.LblMsg = "Verificando conexão de internet..."
    
    'fazendo verificação de conexão com a internet
    If IsWebConnected = True Then
        FrmUpdate.LblMsg = "Conexão com a internet estabelecida"
    Else
        FrmUpdate.LblMsg = "Conexão com a internet não encontrada"
        FrmUpdate.CmdOutraVez.Enabled = True
        Exit Function
    End If
    
    'inicia o update do sistema (download do arquivo executável ProVendas2004.exe)
    FrmUpdate.LblMsg = "Fazendo update do sistema..."
    'Call Download(App.Path & "\ProVendas2004.chm", "http://www.infodigital.inf.br/downloads/sistemas/update/ProVendas2004.chm")
    Call Download(App.Path & "\ProVendas2004.exe", "http://www.infodigital.inf.br/downloads/sistemas/update/ProVendas2004.exe")
    
    
    'resposta do update
     If VPStrUpdate = "OK" Then
        FrmUpdate.LblMsg = "Update do sistema finalizado com sucesso!" & Chr(13) & "É necessário fechar e abrir novamente o sistema Pró Vendas 2004"
        FrmUpdate.CmdAbortar.Visible = False
        FrmUpdate.CmdOutraVez.Visible = False
        FrmUpdate.CmdFechar.Visible = True
     Else
        FrmUpdate.LblMsg = "Ocorreu um erro durante a transferência do arquivo"
        FrmUpdate.CmdOutraVez.Enabled = True
     End If

End Function

Function Download(LocalArquivo As String, LocalURLArquivo As String) As Boolean
     On Error GoTo Baixa_erro
     Dim bt() As Byte

     Open LocalArquivo For Binary Access Write As #1

     bt() = Inet.OpenURL(LocalURLArquivo, icByteArray)

     Put #1, , bt()
     Close #1

     Download = True

     VPStrUpdate = "OK"

     Exit Function

Baixa_erro:

     Download = False

     VPStrUpdate = ""

     Close #1
End Function

