VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmIdentifica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identificação ao sistema"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmIdentifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Fechar sistema"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton CmdAcessar 
      Caption         =   "&Acessar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Acessar sistema"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox TxtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Senha de acesso ao sistema"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox TxtLogin 
      Height          =   285
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "Login de acesso ao sistema"
      Top             =   240
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1200
      OleObjectBlob   =   "FrmIdentifica.frx":000C
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   2040
      OleObjectBlob   =   "FrmIdentifica.frx":0240
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   2040
      OleObjectBlob   =   "FrmIdentifica.frx":02A4
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox PicChave 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      Picture         =   "FrmIdentifica.frx":0308
      ScaleHeight     =   1695
      ScaleWidth      =   1455
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrAlerta As String
Public VPIntAlertaCod As Integer
Public VPStrMsg As String

Private Sub CmdAcessar_Click()
    Screen.MousePointer = vbHourglass

    If TxtLogin.Text = "" Or TxtSenha.Text = "" Then
        VPStrBox = MsgBox("Não pode conter campo em branco.", vbCritical, "Pró Vendas 2004 - Aviso de erro")
        
        If TxtLogin.Text = "" Then
            TxtLogin.SetFocus
        Else
            TxtSenha.SetFocus
        End If
    Else
        
        Conecta

        Dim RecVerif As ADODB.Recordset

        StrSql = "Select * From tb_acesso where Login='" & TxtLogin.Text & "' and Senha='" & TxtSenha.Text & "'"
        Set RecVerif = vgCon.Execute(StrSql)

        If RecVerif.EOF Then    'não achou nenhum registro igual
            VPStrBox = MsgBox("Não foi possível fazer sua identificação no sistema." & Chr(13) & "Verifique se o login e senha digitados estão corretos.", vbExclamation, "Pró Vendas 2004 - Atenção")

        Else
            'CodIdent do usuário
            VGIntIdentCod = RecVerif!CodUsu
            
            'Nome do usuário
            VGStrIdentNome = RecVerif!nome

            'Menus que o usuário poderá acessar
            VGStrMnuVendaCons = RecVerif!VendaCons
            VGStrMnuVendaInc = RecVerif!vendainc
            VGStrMnuVendaExc = RecVerif!vendaexc
            VGStrMnuVendaImp = RecVerif!vendaimp
            VGStrMnuCliCons = RecVerif!clicons
            VGStrMnuCliInc = RecVerif!cliinc
            VGStrMnuCliAlt = RecVerif!clialt
            VGStrMnuCliExc = RecVerif!cliexc
            VGStrMnuCliImp = RecVerif!cliimp
            VGStrMnuFornCons = RecVerif!forncons
            VGStrMnuFornInc = RecVerif!forninc
            VGStrMnuFornAlt = RecVerif!fornalt
            VGStrMnuFornExc = RecVerif!fornexc
            VGStrMnuFornImp = RecVerif!fornimp
            VGStrMnuProdCons = RecVerif!prodcons
            VGStrMnuProdInc = RecVerif!prodinc
            VGStrMnuProdAlt = RecVerif!prodalt
            VGStrMnuProdExc = RecVerif!prodexc
            VGStrMnuProdImp = RecVerif!prodimp
            VGStrMnuEstCons = RecVerif!estcons
            VGStrMnuEstInc = RecVerif!estinc
            VGStrMnuEstAlt = RecVerif!estalt
            VGStrMnuEstExc = RecVerif!estexc
            VGStrMnuEstImp = RecVerif!estimp
            VGStrMnuEstAlerta = RecVerif!estalerta
            VGStrMnuCredCons = RecVerif!credcons
            VGStrMnuCredExc = RecVerif!credexc
            VGStrMnuCredImp = RecVerif!credimp
            VGStrMnuParcAlt = RecVerif!parcalt
            VGStrMnuParcExc = RecVerif!parcexc
            VGStrMnuCredstaCons = RecVerif!credstacons
            VGStrMnuCredstaInc = RecVerif!credstainc
            VGStrMnuCredstaAlt = RecVerif!credstaalt
            VGStrMnuCredstaExc = RecVerif!credstaexc
            VGStrMnuCredstaImp = RecVerif!credstaimp
            VGStrMnuCxCons = RecVerif!cxcons
            VGStrMnuCxInc = RecVerif!cxinc
            VGStrMnuCxAlt = RecVerif!cxalt
            VGStrMnuCxExc = RecVerif!cxexc
            VGStrMnuCxImp = RecVerif!cximp
            VGStrMnuPagCons = RecVerif!pagcons
            VGStrMnuPagInc = RecVerif!paginc
            VGStrMnuPagAlt = RecVerif!pagalt
            VGStrMnuPagExc = RecVerif!pagexc
            VGStrMnuPagImp = RecVerif!pagimp
            VGStrMnuPagBx = RecVerif!pagbx
            VGStrMnuRecCons = RecVerif!pagcons
            VGStrMnuRecInc = RecVerif!recinc
            VGStrMnuRecAlt = RecVerif!recalt
            VGStrMnuRecExc = RecVerif!recexc
            VGStrMnuRecImp = RecVerif!recimp
            VGStrMnuRecBx = RecVerif!RecBx
            VGStrMnuOrcCons = RecVerif!orccons
            VGStrMnuOrcInc = RecVerif!orcinc
            VGStrMnuOrcAlt = RecVerif!orcalt
            VGStrMnuOrcExc = RecVerif!orcexc
            VGStrMnuOrcImp = RecVerif!orcimp
            VGStrMnuVendCons = RecVerif!vendcons
            VGStrMnuVendInc = RecVerif!vencinc
            VGStrMnuVendAlt = RecVerif!vendalt
            VGStrMnuVendExc = RecVerif!vendexc
            VGStrMnuVendImp = RecVerif!vendimp
            VGStrMnuExNiver = RecVerif!exniver
            VGStrMnuExCob = RecVerif!excob
            VGStrMnuExMala = RecVerif!exmala
            VGStrMnuPropImp = RecVerif!propimp
            VGStrMnuCarneImp = RecVerif!carneimp

            Unload Me

            MDIPrincipal.Caption = "Pró Vendas 2004 - Sistema Integrado de Vendas                                                 Usuário: " & VGStrIdentNome
         End If

        Desconecta

    End If

    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdFechar_Click()
    Screen.MousePointer = vbHourglass
    
    If VGStrAcesso = "acesso" Then
        Unload Me
    ElseIf VGStrAcesso = "" Then
        Unload Me
        Unload MDIPrincipal
    End If
    
    VGStrAcesso = ""
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Height = 2445
    Width = 4755
        
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmIdentifica.Left = (MDIPrincipal.Width / 2) - (FrmIdentifica.Width / 2)
  FrmIdentifica.Top = (MDIPrincipal.Height / 3) - (FrmIdentifica.Height / 3)
End Sub

Private Sub TxtLogin_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        TxtSenha.SetFocus
    End If

    If KeyAscii >= 1 And KeyAscii <= 7 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 9 And KeyAscii <= 29 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 32 And KeyAscii <= 47 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 58 And KeyAscii <= 64 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 91 And KeyAscii <= 96 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 123 And KeyAscii <= 127 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 192 And KeyAscii <= 196 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 200 And KeyAscii <= 207 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 210 And KeyAscii <= 214 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 217 And KeyAscii <= 221 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 224 And KeyAscii <= 228 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 232 And KeyAscii <= 239 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 242 And KeyAscii <= 246 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 249 And KeyAscii <= 253 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdAcessar.SetFocus
    End If
        
    If KeyAscii >= 1 And KeyAscii <= 7 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 9 And KeyAscii <= 29 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 32 And KeyAscii <= 47 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 58 And KeyAscii <= 64 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 91 And KeyAscii <= 96 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 123 And KeyAscii <= 127 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 192 And KeyAscii <= 196 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 200 And KeyAscii <= 207 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 210 And KeyAscii <= 214 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 217 And KeyAscii <= 221 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 224 And KeyAscii <= 228 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 232 And KeyAscii <= 239 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 242 And KeyAscii <= 246 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 249 And KeyAscii <= 253 Then
        KeyAscii = 0
    End If

End Sub
