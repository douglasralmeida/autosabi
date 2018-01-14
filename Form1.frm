VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frbRequerimentosdoDia 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   4155
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pctFundo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3825
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox chkImpressãoautomática 
         Caption         =   "Impressão automática"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   2280
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox pctFundoProgresso 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   6
         Top             =   1920
         Width           =   975
         Begin VB.PictureBox pctProgresso 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   0
            ScaleHeight     =   465
            ScaleWidth      =   225
            TabIndex        =   7
            Top             =   0
            Width           =   255
            Begin VB.Label lblProgresso2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Progresso"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Left            =   0
               TabIndex        =   9
               Top             =   -25
               Width           =   30000
            End
         End
         Begin VB.Label lblProgresso 
            BackStyle       =   0  'Transparent
            Caption         =   "Progresso"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   20
            TabIndex        =   8
            Top             =   -15
            Width           =   30000
         End
      End
      Begin VB.Timer tmImprimirMarcação 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1200
         Top             =   480
      End
      Begin VB.ListBox lstMostrarRequerimentos 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   480
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ListBox lstClassificar 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer tmRelaçãodeRequerimentos 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   720
         Top             =   480
      End
      Begin VB.Timer tmVerificaseMenuRequerimentosfoiAcionado 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   1680
      End
      Begin VB.Timer tmVeriricaSeControleEstaAberto 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   120
         Top             =   480
      End
      Begin RichTextLib.RichTextBox rtbRequerimentos 
         Height          =   615
         Left            =   1440
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form1.frx":08CA
      End
      Begin VB.Label cmdImprimir 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " 2ª Via Marcação de Exame "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   240
         Left            =   600
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblRequerimentodoSABI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Requerimentos do SABI"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   80
         Width           =   3615
      End
      Begin VB.Label lblversão 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vilton 5.07.16 P"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Image imgFecha 
         Height          =   480
         Left            =   2280
         Picture         =   "Form1.frx":0959
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblImprimir 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " Imprimir  "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   240
         Left            =   300
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "frbRequerimentosdoDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
    End Type

    
    Dim MenuName As New Collection
    Dim MenuHandle As New Collection
    Dim lHwnd As Long
    Dim imprimereq As Boolean
    Dim modoImprime As String
    
    
    Private Const SW_MINIMIZE = 6, SW_NORMAL = 1, SW_MAXIMIZE = 3, SW_RESTORE = 9
    
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

   Private Const MK_LBUTTON = &H1

   Private Const BM_CLICK = &HF5
    Private Const WM_CLOSE = &H10
    Private Const WM_COMMAND = &H111
    'Private Const WM_LBUTTONUP = &H202
    Private Const BN_CLICKED = 0
     Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Const MOUSEEVENTF_MIDDLEUP = &H40
    Const MOUSEEVENTF_MOVE = &H1
    Const MOUSEEVENTF_ABSOLUTE = &H8000
    Const MOUSEEVENTF_RIGHTDOWN = &H8
    Const MOUSEEVENTF_RIGHTUP = &H10
    Const WM_SETTEXT As Long = &HC
    Const WM_GETTEXTLENGTH = &HE
    Const WM_GETTEXT As Integer = &HD
    Const LB_SETSEL = &H185
    Const CB_SETCURSEL = &H14E
    Const WM_KEYDOWN = &H100
    Const VK_RETURN = &HD
    Const RDW_INVALIDATE = 1
    Const VK_LBUTTON = &H1


Private Const GW_HWNDNEXT = 2

    Private Const MF_BYPOSITION = &H400&



    Private Declare Function apiRedrawWindow Lib "USER32" Alias "RedrawWindow" (ByVal hwnd As Long, ByVal lprcUpdate As Boolean, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
    Private Declare Function SendMessageString Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Private Declare Function GetClassName Lib "USER32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function SendMessage2 Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
    
    Private Declare Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
    Private Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer
    Private Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function GetWindowRect Lib "USER32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetMenu Lib "USER32" (ByVal hwnd As Long) As Long
    Private Declare Function GetMenuItemCount Lib "USER32" (ByVal hMenu As Long) As Long
    Private Declare Function GetSubMenu Lib "USER32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
    Private Declare Function GetMenuString Lib "USER32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, _
    ByVal cch As Long) As Long
    Private Declare Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindow Lib "USER32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Private Declare Function GetWindowPlacement Lib "USER32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
    Private Declare Function SetWindowPlacement Lib "USER32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
    Private Declare Function FindWindowA Lib "USER32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Sub ColocaTelaControleOperacionanoModoNormal()
    If GlobalIDControleOperacional Then
        If AppToForeground(, GlobalIDControleOperacional, SW_NORMAL) Then
           ' a tela foi maximizada
        Else
            MsgBox "Falha ao maximizar a tela Controle Operacional", vbCritical, "Maximizar Controle Operacional"
        End If
    Else
        MsgBox "Não foi encontrada tela Controle Operaciona", vbCritical, "Maximizar Controle Operacional"
    End If

End Sub
Sub atualizaprogresso()
    
    Dim memo As String
    pctFundoProgresso.Visible = True
    memo = CDate(Time - GlobalHoradeInicio) & " > " & Format(GlobalIDRequerimento, "000") & "/" & Format(GlobalQuantidadedeRequerimentos, "000")
    lblProgresso.Caption = memo & " " & GlobalRequerimentos(GlobalIDRequerimento).Segurado
    lblProgresso2.Caption = memo & " " & GlobalRequerimentos(GlobalIDRequerimento).Segurado
    pctProgresso.Width = pctFundoProgresso.Width * (GlobalIDRequerimento / GlobalQuantidadedeRequerimentos)
    DoEvents

End Sub
Sub verificaeapaga(tituladatela As String)
    Dim IDTelaExterna As Long
    IDTelaExterna = 0
    IDTelaExterna = ObtemTelaPrincipalporTitulo(tituladatela)
    If IDTelaExterna <> 0 Then SendMessage IDTelaExterna, WM_CLOSE, 0, 0

End Sub
    Function abrePesquisaAvançada() As String
        Dim contavezes As Long
        ColocaTelaControleOperacionanoModoMaximizado
        DoEvents
        MouseCliqueAbsoluto 690, 150
        Sleep 1000
                
        abrePesquisaAvançada = "A tela 'Pesquisa Avançada' não abriu."
        GlobalIDTelaPesquisaAvançada = 0
        contavezes = 0
        While GlobalIDTelaPesquisaAvançada = 0
            GlobalIDTelaPesquisaAvançada = ObtemTelaPrincipalporTitulo("Pesquisa Avançada")
            Sleep 300
            DoEvents
            contavezes = contavezes + 1
            If contavezes > 20 Then Exit Function
            
        Wend
        abrePesquisaAvançada = GlobalIDTelaPesquisaAvançada
    
    End Function
    Function AbreSegundaViaMarcaçãodeExame(numerodonit As String) As String
        Dim ValorNit As String
        Dim contaloop As Long
        Dim hThunderRT6FrameIMPRIME  As Long
        Dim hImMaskWndCIassIMPRIME As Long
        Dim hThunderRT6FormDC As Long
        Dim hThunderRT6CommandButtonVisualizar As Long
        Dim hThunderRT6CommandButtonCancelar As Long
        Dim hThunderRT6CommandButtonImprimir As Long
        Dim sizevisualizar As RECT
        Dim sizecancelar As RECT
        Dim sizeimprimir As RECT
        Dim hMDIClient As Long
        Dim strClass As String
        Dim sStr As String
        Dim nret As String
        Dim res As String
        Dim size As RECT
        Dim lParam As Long
        Dim cursortecla As POINTAPI
        Dim botaoimprimirenabled As Boolean
        GlobalTeclaMarcaçãdeExame = ""
        AbreSegundaViaMarcaçãodeExame = "Falha na Segunda Via de Marcação de Exame"

        ValorNit = numerodonit
        
        'abre tela de impresssao
        While InStr(1, ValorNit, ".")
            ValorNit = Mid(ValorNit, 1, InStr(1, ValorNit, ".") - 1) & Mid(ValorNit, InStr(1, ValorNit, ".") + 1)
        Wend
        'ValorNIT = Mid(ValorNIT, 1, Len(ValorNIT) - 2)
        ClickMenu GlobalIDControleOperacional, 4, 7
        DoEvents
        Sleep 100
        hMDIClient = 0
        While hMDIClient = 0
            hMDIClient = FindWindowEx(GlobalIDControleOperacional, 0, "MDIClient", "")
            DoEvents
            Sleep 100
        Wend
        hThunderRT6FormDC = 0
        While hThunderRT6FormDC = 0
            hThunderRT6FormDC = FindWindowEx(hMDIClient, 0, "ThunderRT6FormDC", "Segunda Via de Marcação de Exame")
            DoEvents
            Sleep 100
        Wend
        GlobalIDTelaSegundaVia = hThunderRT6FormDC
        res = SetWindowPos(GlobalIDTelaSegundaVia, 0, 0, 0, 800, 200, 0)
        DoEvents
        'encontra a tecla Visualizar e desloca para fora da area visivel
        hThunderRT6CommandButtonVisualizar = FindWindowEx(GlobalIDTelaSegundaVia, 0, "ThunderRT6CommandButton", "&Visualizar")
        hThunderRT6CommandButtonCancelar = FindWindowEx(GlobalIDTelaSegundaVia, 0, "ThunderRT6CommandButton", "&Cancelar")
        hThunderRT6CommandButtonImprimir = FindWindowEx(GlobalIDTelaSegundaVia, 0, "ThunderRT6CommandButton", "&Imprimir")
        'não tira mais a tecla visualizar
        'If hThunderRT6CommandButtonVisualizar <> 0 Then SetWindowPos hThunderRT6CommandButtonVisualizar, 0, 1000, 1000, 10, 10, 0
        'encontra o campo de NIT
        hThunderRT6FrameIMPRIME = FindWindowEx(GlobalIDTelaSegundaVia, 0, "ThunderRT6Frame", "NIT Requerente")
        If hThunderRT6FrameIMPRIME <> 0 Then
            hImMaskWndCIassIMPRIME = FindWindowEx(hThunderRT6FrameIMPRIME, 0, vbNullString, vbNullString)
            strClass = Space(100)
            nret = GetClassName(hImMaskWndCIassIMPRIME, strClass, 100)
            If Mid(strClass, 1, 14) = "ImMaskWndClass" Then
            
                res = GetWindowRect(hImMaskWndCIassIMPRIME, size)
                lParam = convlong(size.Left + (size.Right - size.Left) / 2, size.Top + (size.Bottom - size.Top) / 2)
                    
                SendMessage hImMaskWndCIassIMPRIME, WM_LBUTTONDOWN, MK_LBUTTON, lParam
                SendMessage hImMaskWndCIassIMPRIME, WM_LBUTTONUP, MK_LBUTTON, lParam
                DoEvents
                Sleep 100
                SendKeys Mid(ValorNit, 1, Len(ValorNit) - 2)
                DoEvents
                Sleep 100
                SendKeys Right$(ValorNit, 1)
                DoEvents
                Sleep 100
                If WindowTextGet(hImMaskWndCIassIMPRIME) <> ValorNit Then
                    AbreSegundaViaMarcaçãodeExame = WindowTextGet(hImMaskWndCIassIMPRIME) & " diferente de " & ValorNit
                    Exit Function
                End If
            
                
                
                botaoimprimirenabled = Val(GetWindowLong(hThunderRT6CommandButtonImprimir, GWL_STYLE) And WS_DISABLED) = 0
                contaloop = 0
                While botaoimprimirenabled = False
                    botaoimprimirenabled = Val(GetWindowLong(hThunderRT6CommandButtonImprimir, GWL_STYLE) And WS_DISABLED) = 0
                    Sleep 100
                    contaloop = contaloop + 1
                    If contaloop > 100 Then
                        AbreSegundaViaMarcaçãodeExame = "O SABI não retornou em 10 segundos os dados do requerimento para o NIT '" & ValorNit & "' informado."
                        Exit Function
                    End If
                Wend
                Beep
                If chkImpressãoautomática.Value = False Then
                
                    While GetAsyncKeyState(VK_LBUTTON) = True
                    Wend '
                    While GetAsyncKeyState(VK_LBUTTON) = False
                    Wend '
                    While GetAsyncKeyState(VK_LBUTTON) = True
                    Wend '
                    GetCursorPos cursortecla
                    res = GetWindowRect(hThunderRT6CommandButtonVisualizar, sizevisualizar)
                    res = GetWindowRect(hThunderRT6CommandButtonCancelar, sizecancelar)
                    res = GetWindowRect(hThunderRT6CommandButtonImprimir, sizeimprimir)
    
                    If cursortecla.Y >= sizevisualizar.Top And cursortecla.Y <= sizevisualizar.Bottom Then
                        'mouse clicado na linha das teclas
                        If cursortecla.X >= sizevisualizar.Left And cursortecla.X <= sizevisualizar.Right Then
                            GlobalTeclaMarcaçãdeExame = "Visualizar"
                        End If
                        If cursortecla.X >= sizeimprimir.Left And cursortecla.X <= sizeimprimir.Right Then
                            GlobalTeclaMarcaçãdeExame = "Imprimir"
                        End If
                        If cursortecla.X >= sizecancelar.Left And cursortecla.X <= sizecancelar.Right Then
                            GlobalTeclaMarcaçãdeExame = "Cancelar"
                        End If
                        
                    
                    End If
                    If GlobalTeclaMarcaçãdeExame = "Visualizar" Or GlobalTeclaMarcaçãdeExame = "Imprimir" Then
                        GlobalRelaçaodeRequerimentosImpressos = GlobalRelaçaodeRequerimentosImpressos & ">" & GlobalRequerimentos(GlobalIDRequerimento).Número & "<"
                        If modoImprime = "Todos" Then atualizalista
                    End If
                    If GlobalTeclaMarcaçãdeExame = "Visualizar" Then
                        GlobalIDTelaMarcaçãodeExameCrystalReport = ProcuraCrystal
                        While GlobalIDTelaMarcaçãodeExameCrystalReport = 0
                            GlobalIDTelaMarcaçãodeExameCrystalReport = ProcuraCrystal
                            Sleep 100
                            DoEvents
                        Wend
                        res = SetWindowPos(GlobalIDTelaMarcaçãodeExameCrystalReport, 0, Screen.Width / 30 - 100, 0, Screen.Width / 30 + 100, (Screen.Height - 600) / 15, 0)
                        res = SetWindowPos(GlobalIDTelaMarcaçãodeExameCrystalReport, -1, 0, 0, 0, 0, 3)
                        SetForegroundWindow (GlobalIDTelaMarcaçãodeExameCrystalReport)
                        'espera a tela ser fechada
                        While GlobalIDTelaMarcaçãodeExameCrystalReport <> 0
                            GlobalIDTelaMarcaçãodeExameCrystalReport = ProcuraCrystal
                            Sleep 300
                            DoEvents
                        Wend
                        AbreSegundaViaMarcaçãodeExame = ""
                        Exit Function
                    Else
                        'espera a tela ser fechada
                        While hThunderRT6FormDC <> 0
                            hThunderRT6FormDC = FindWindowEx(hMDIClient, 0, "ThunderRT6FormDC", "Segunda Via de Marcação de Exame")
                            DoEvents
                            Sleep 300
                        Wend
            
                        AbreSegundaViaMarcaçãodeExame = ""
                        Exit Function
    
                    End If
                Else
                    Sleep 100
                    SendMessage hThunderRT6CommandButtonImprimir, BM_CLICK, 0, 0
                    DoEvents
                    GlobalRelaçaodeRequerimentosImpressos = GlobalRelaçaodeRequerimentosImpressos & ">" & GlobalRequerimentos(GlobalIDRequerimento).Número & "<"
                    If modoImprime = "Todos" Then atualizalista
                    'espera a tela ser fechada
                    While hThunderRT6FormDC <> 0
                        hThunderRT6FormDC = FindWindowEx(hMDIClient, 0, "ThunderRT6FormDC", "Segunda Via de Marcação de Exame")
                        DoEvents
                        Sleep 300
                    Wend
        
                    AbreSegundaViaMarcaçãodeExame = ""
                    Exit Function

                End If
            End If
        End If
    End Function
    Function ExtraiNIT() As String
           Dim contaloop As Long
           Dim hPrimeira As Long
           Dim ClassedoControle As String
           Dim classlength As Long
           Dim hBotaoSairNIT As Long
           Dim hValorNIT As Long
           Dim ValorNit As String
           ExtraiNIT = "0.000.000.000-0"
            'verifica se tela Informações de NIT(s) Secundário(s)' esta aberta
            contaloop = 0
            GlobalIDTelaNITSecundario = 0
            While GlobalIDTelaNITSecundario = 0 And contaloop < 50
                GlobalIDTelaNITSecundario = ObtemTelaPrincipalporTitulo("Informações de NIT(s) Secundário(s)")
                Sleep 300
                DoEvents
                GlobalIDTelaConsultaSemCriterio = ObtemTelaPrincipalporTitulo("AVISO IMPORTANTE")
                If GlobalIDTelaConsultaSemCriterio <> 0 Then
                    MsgBox "CONSULTA SEM CRITERIO"
                    Exit Function
                End If
                contaloop = contaloop + 1
                'NO dia 08/07/2015 APS/DIVINÓPOLIS aparece entre os requerimento
                '15:20        EDGAR RIBEIRO  DE OLIVEIRA  MORAES           N       166101916
                'APS -  BELO HORIZONTE-OESTE
                'este loop fica rodando sem fim
            Wend
            If contaloop = 50 Then
                MsgBox "A tela de NIT não abriu."
                Exit Function
            End If
                
            hValorNIT = 0
            While hValorNIT = 0 And contaloop < 50
                hPrimeira = FindWindowEx(GlobalIDTelaNITSecundario, 0, vbNullString, vbNullString)
                Do While hPrimeira <> 0 And contaloop < 50
                    ClassedoControle = Space(128)
                    classlength = GetClassName(hPrimeira, ClassedoControle, 128)
                    ClassedoControle = Left(ClassedoControle, classlength)
                    If ClassedoControle = "ThunderRT6CommandButton" Then hBotaoSairNIT = hPrimeira
                    If ClassedoControle = "ImMaskWndClass" Then hValorNIT = hPrimeira
                    hPrimeira = FindWindowEx(GlobalIDTelaNITSecundario, hPrimeira, vbNullString, vbNullString)
                    'sleep  300
                    DoEvents
                    contaloop = contaloop + 1
                Loop
            Wend
            Sleep 300
            ValorNit = WindowTextGet(hValorNIT)
            ExtraiNIT = ValorNit
            'sleep  300
        'Else
            'ValorNIT = memoerro
        'End If
        SendMessage hBotaoSairNIT, BM_CLICK, 0, 0
        Sleep 300
        SendMessage GlobalIDTelaNITSecundario, WM_CLOSE, 0, 0
    End Function

    Sub AbrirTelaPesquisaAvançada(numerodorequerimento As String)
        Dim contaloop As Long
        MsgBox "Clique em Avançado e quando aparecer a tela clique em OK", vbApplicationModal, "Pesquisa Avançada"
        contaloop = 0
        'GlobalIDTelaPesquisaAvançada = GetForegroundWindow
        While GlobalIDTelaPesquisaAvançada = 0 And contaloop < 10
            GlobalIDTelaPesquisaAvançada = ObtemTelaPrincipalporTitulo("Pesquisa Avançada")
            Sleep 300
            DoEvents
            contaloop = contaloop + 1
        Wend
        res = SetWindowPos(GlobalIDTelaPesquisaAvançada, 0, 0, 0, 800, 460, 0)
        
        'Clipboard.SetText numerodorequerimento
        'MsgBox "Cole o numero do requerimento e Clique OK "
        While GetAsyncKeyState(VK_LBUTTON) = True
        Wend '

        While GetAsyncKeyState(VK_LBUTTON) = False
        Wend '
        While GetAsyncKeyState(VK_LBUTTON) = True
        Wend '

        Sleep 300
        SendKeys numerodorequerimento
        DoEvents
        Sleep 200
        SendKeys "{BS}"
        DoEvents
        Sleep 200
        SendKeys Right$(numerodorequerimento, 1)
        DoEvents
        Sleep 600
       MouseClique 565, 342
        Sleep 600
        SendMessage GlobalIDTelaPesquisaAvançada, WM_CLOSE, 0, 0

        '------------

    End Sub
    
    Sub ConverteData(datalonga As String)
    Dim memo As String
    Dim dia As Long
    Dim mes As Long
    Dim ano As Long
    GlobalDatadosRequerimentos = "00000000"
    memo = UCase(datalonga)
    
    If InStr(1, memo, ",") Then memo = Trim(Mid(memo, InStr(1, memo, ",") + 1))
    dia = Val(memo)
    ano = Val(Right$(memo, 4))
    If InStr(1, memo, "DE") Then memo = Trim(Mid(memo, InStr(1, memo, "DE") + 3))
    If InStr(1, memo, "DE") Then memo = Trim(Mid(memo, 1, InStr(1, memo, "DE") - 1))
    Select Case memo
        Case "JANEIRO"
        mes = 1
        Case "FEVEREIRO"
        mes = 2
        Case "MARÇO"
        mes = 3
        Case "ABRIL"
        mes = 4
        Case "MAIO"
        mes = 5
        Case "JUNHO"
        mes = 6
        Case "JULHO"
        mes = 7
        Case "AGOSTO"
        mes = 8
        Case "SETEMBRO"
        mes = 9
        Case "OUTUBRO"
        mes = 10
        Case "NOVEMBRO"
        mes = 11
        Case "DEZEMBRO"
        mes = 12
    End Select
    GlobalDatadosRequerimentos = Format(ano, "0000") & Format(mes, "00") & Format(dia, "00")

End Sub

     Function convlong(xini As Long, yini As Long)
        Dim lParam As Long
        lParam = 256 * 64
        lParam = yini * (lParam * 4) + xini
        convlong = lParam
     End Function
    
Sub Clicka(tela As Long, X As Long, Y As Long)
    Dim PHANDLE As Long
    Dim lParam As Long
    lParam = convlong(X, Y)
    PHANDLE = tela
    SendMessage PHANDLE, WM_LBUTTONDOWN, MK_LBUTTON, lParam
    SendMessage PHANDLE, WM_LBUTTONUP, 0, lParam
End Sub

Public Function GET_Y_LPARAM(ByVal lParam As Long) As Long
  Dim HexStr As String
  HexStr = Right("00000000" & Hex(lParam), 8)
  GET_Y_LPARAM = CLng("&H" & Left(HexStr, 4))
End Function

Public Function GET_X_LPARAM(ByVal lParam As Long) As Long
  Dim HexStr As String
  HexStr = Right("00000000" & Hex(lParam), 8)
  GET_X_LPARAM = CLng("&H" & Right(HexStr, 4))
End Function
    Function DialogGetHwnd(Optional ByVal sDialogCaption As String = vbNullString, Optional sClassName As String = vbNullString) As Long
    On Error Resume Next
    DialogGetHwnd = FindWindowA(sClassName, sDialogCaption)
    On Error GoTo 0
End Function

Function AppToForeground(Optional sFormCaption As String, Optional lHwnd As Long, Optional lWindowState As Long = SW_NORMAL) As Boolean
    Dim tWinPlace As WINDOWPLACEMENT

    If lHwnd = 0 Then
        lHwnd = DialogGetHwnd(sFormCaption)
    End If
    If lHwnd Then
        tWinPlace.Length = Len(tWinPlace)
        'Get the windows current placement
        Call GetWindowPlacement(lHwnd, tWinPlace)
        'Set the windows placement
        tWinPlace.showCmd = lWindowState
        'Change window state
        Call SetWindowPlacement(lHwnd, tWinPlace)
        'Bring to foreground
        AppToForeground = SetForegroundWindow(lHwnd)
    End If
End Function

Sub ColocaTelaControleOperacionanoModoMaximizado()
    If GlobalIDControleOperacional Then
        If AppToForeground(, GlobalIDControleOperacional, SW_MAXIMIZE) Then
           ' a tela foi maximizada
        Else
            MsgBox "Falha ao maximizar a tela Controle Operacional", vbCritical, "Maximiza Controle Operacional"
        End If
    Else
        MsgBox "Não foi encontrada tela Controle Operacional"
    End If

End Sub

    Sub MouseCliqueAbsoluto(posx As Long, posy As Long)
        Dim size As RECT
        Dim pt As POINTAPI
        GetCursorPos pt
        SetCursorPos posx, posy
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        SetCursorPos pt.X, pt.Y
    End Sub

    Sub atualizalista()
        On Error Resume Next
        Dim conta As Long
        
        lstMostrarRequerimentos.Visible = False
        lstMostrarRequerimentos.Clear
        lstMostrarRequerimentos.AddItem "Seq. Requerimento NIT" & Chr(9) & Chr(9) & "Segurado"
        For conta = 1 To GlobalQuantidadedeRequerimentos
            If GlobalRequerimentos(conta).NIT <> "" Then
                lstMostrarRequerimentos.AddItem Format(conta, "000") & "  " & GlobalRequerimentos(conta).Número & Chr(9) & GlobalRequerimentos(conta).NIT & Chr(9) & GlobalRequerimentos(conta).Segurado
            Else
                lstMostrarRequerimentos.AddItem Format(conta, "000") & "  " & GlobalRequerimentos(conta).Número & Chr(9) & Chr(9) & Chr(9) & GlobalRequerimentos(conta).Segurado
            End If
            lstMostrarRequerimentos.Selected(conta) = InStr(1, GlobalRelaçaodeRequerimentosImpressos, ">" & GlobalRequerimentos(conta).Número & "<")
        Next conta
        lstMostrarRequerimentos.Visible = True
        DoEvents
    
    End Sub
    Private Function WindowTextGet(ByVal hwnd As Long) As String
    Dim strBuff As String, lngLen As Long
    lngLen = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
    If lngLen > 0 Then
        lngLen = lngLen + 1
        strBuff = String(lngLen, vbNullChar)
        lngLen = SendMessage(hwnd, WM_GETTEXT, lngLen, ByVal strBuff)
        WindowTextGet = Left(strBuff, lngLen)
    End If
End Function

Function ProcuraCrystal() As Long
    Dim lngHWnd As Long
    Dim lngHWnd2 As Long
    Dim titletmp As String
    Dim nret As Long
    Dim size As RECT
    Dim lhWndP As Long
    ProcuraCrystal = 0
    lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
    Do While lhWndP <> 0
        If Crystal(lhWndP) Then
            ProcuraCrystal = lhWndP 'tela externa
            Exit Function
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop

End Function

Private Sub ClickOpen(hMsgBox As Long)
    Dim hButtonOpen As Long
    Dim hComboBox As Long
    hButtonOpen = FindWindowEx(hMsgBox, 0, "Button", "OK")
    'hComboBox = FindWindowEx(hMsgBox, 0, "ComboBox", "")
    'If hButtonOpen = 0 Then Stop
    SendMessage hButtonOpen, BM_CLICK, 0, 0
End Sub
Private Function Crystal(hMsgBox As Long) As Boolean
    Dim hButtonOpen As Long
    Dim nret As Long
    Dim res As String
    Dim size As RECT
    Dim larguratotal As Long
    Crystal = False
    res = GetWindowRect(hMsgBox, size)
    larguratotal = size.Right - size.Left
    hButtonOpen = FindWindowEx(hMsgBox, 0, "ToolbarWindow32", "")
    If hButtonOpen > 0 Then
        res = GetWindowRect(hButtonOpen, size)
        If larguratotal - (size.Right - size.Left) = 16 And size.Bottom - size.Top = 28 Then
            Crystal = True
            Exit Function
        End If
    End If
End Function
    Sub MouseCliquePara(posx As Long, posy As Long)
        Dim size As RECT
        Dim IDTelaAtiva As Long
        Dim pt As POINTAPI
        GetCursorPos pt
        IDTelaAtiva = GetForegroundWindow
        res = GetWindowRect(IDTelaAtiva, size)
        SetCursorPos size.Left + posx, size.Top + posy
        Sleep 10000
        SetCursorPos pt.X, pt.Y
    End Sub
Function achaTelaInternaAtiva(NomedaTela As String) As Long
    Dim lngHWnd As Long
    Dim lngHWnd2 As Long
    Dim titletmp As String
    Dim nret As Long
    Dim size As RECT
    Dim lhWndP As Long
    achaTelaInternaAtiva = 0
    lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
    Do While lhWndP <> 0
    '-------------------
    lngHWnd = FindWindowEx(lhWndP, 0, vbNullString, vbNullString)
    Do Until lngHWnd = 0
        If lngHWnd > 0 Then
            lngHWnd2 = FindWindowEx(lngHWnd, 0, vbNullString, vbNullString)
            Do Until lngHWnd2 = 0
                If lngHWnd2 > 0 Then
                    titletmp = Space(256)
                    nret = GetWindowText(lngHWnd2, titletmp, Len(titletmp))
                    If InStr(1, titletmp, NomedaTela) > 0 Then
                        achaTelaInternaAtiva = lngHWnd2
                        Exit Function
                    End If
                End If
                lngHWnd2 = FindWindowEx(lngHWnd2, lngHWnd, vbNullString, vbNullString)
            Loop
        End If
        lngHWnd = FindWindowEx(lhWndP, lngHWnd, vbNullString, vbNullString)
    Loop

    
    '---------------
    
    
    
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop

End Function

Function esperaCRYSTALREPORTeExporta() As String
    Dim COsize As RECT
    Dim size As RECT
    Dim titletmp As String
    Dim nret As Long
    Dim TelaSize As RECT
    Dim arquivo As String
    Dim memo As String
    Dim hDC As Long
    Dim lcount As Long
    Dim LocalIDBotãoSalvar As Long
    hDC = GetWindowDC(0)
    Dim hNomedoArquivo As Long
    Dim hDUIView As Long
    Dim hDirectUI  As Long
    Dim hFloatNotify As Long
    Dim hComboBox As Long
    Dim hBotãoSalvar As Long
    Dim hDestinoExport As Long
    Dim hBotãoOKExport As Long
    Dim hFormatoExport As Long
    Dim conta As Long

    GlobalIDTelaSalvarComo = 0
    tmRelaçãodeRequerimentos.Enabled = False

    'esperaCRYSTALREPORTeExporta
    GlobalIDTelaRequerimentosCrystalReport = ProcuraCrystal
    While GlobalIDTelaRequerimentosCrystalReport = 0
        GlobalIDTelaRequerimentosCrystalReport = ProcuraCrystal
        Sleep 300
        DoEvents
    Wend
    res = SetWindowPos(GlobalIDTelaRequerimentosCrystalReport, 0, 0, 0, 800, 460, 0)
    SetForegroundWindow (GlobalIDTelaRequerimentosCrystalReport)

    Sleep 1000  'com 300 falhou com o Bruno - clicou antes da hora
    'implementar rotina que repete o clique periodicamente ate´ vir a nova tela
    'SetForegroundWindow (GlobalIDTelaRequerimentosCrystalReport)
    'sleep  1000
    SetForegroundWindow (GlobalIDTelaRequerimentosCrystalReport)
    DoEvents
    Sleep 500
    MouseClique 262, 44
    DoEvents
    Sleep 500
    '    nao funcionou
                
    'espera tela Export
    titletmp = Space(256)
    nret = GetWindowText(GlobalIDTelaAtiva, titletmp, Len(titletmp))
    GlobalTítulodaTelaAtiva = titletmp
    While Mid(GlobalTítulodaTelaAtiva, 1, 6) <> "Export"
        GlobalIDTelaAtiva = GetForegroundWindow
        titletmp = Space(256)
        nret = GetWindowText(GlobalIDTelaAtiva, titletmp, Len(titletmp))
        GlobalTítulodaTelaAtiva = titletmp
        DoEvents
        Sleep 300
        If InStr(1, titletmp, "SABI - Controle Operacional") > 0 Then
            ClickOpen (GlobalIDTelaAtiva)
        Else
            If Len(Trim(titletmp)) = 1 Then MouseClique 262, 44
        End If
    Wend
    
    'espera a tela export
    GlobalIDTelaExport = ObtemTelaPrincipalporTitulo("Export")
    While GlobalIDTelaExport = 0
        Sleep 300
        DoEvents
        GlobalIDTelaExport = ObtemTelaPrincipalporTitulo("Export")
    Wend
    
    If GlobalIDTelaExport > 0 Then
        lcount = 0
        hDestinoExport = 0
        Do While hDestinoExport = 0 Or lcount > 10
            hBotãoOKExport = FindWindowEx(GlobalIDTelaExport, 0, "Button", "OK")
            'encontra DirectUIHWND
            hFormatoExport = FindWindowEx(GlobalIDTelaExport, 0, "ComboBox", "")
            hDestinoExport = FindWindowEx(GlobalIDTelaExport, hFormatoExport, "ComboBox", "")
            lcount = lcount + 1
            Sleep 300
            DoEvents
        Loop
        'o destino deve ser escolhido antes do formato para não gerar erro de e-mail não configurado
        
        For conta = 0 To 100
            SendMessage hDestinoExport, CB_SETCURSEL, conta, 0&
            DoEvents
            If WindowTextGet(hDestinoExport) = "Disk file" Then Exit For
        Next conta
        If conta > 99 Then
            MsgBox "Não foi encontrado o destino 'Disk file'", vbCritical, "Export"
            End
        End If
        For conta = 0 To 100
            SendMessage hFormatoExport, CB_SETCURSEL, conta, 0&
            DoEvents
            If WindowTextGet(hFormatoExport) = "Rich Text Format" Then Exit For
        Next conta
        If conta > 99 Then
            MsgBox "Não foi encontrado o destino 'Disk file'", vbCritical, "Export"
            End
        End If

        Sleep 100
        SendMessage hBotãoOKExport, BM_CLICK, 0, 0
        DoEvents
        Sleep 300
        DoEvents
        Sleep 1000
        SendMessage GlobalIDTelaExport, WM_CLOSE, 0, 0
        DoEvents
    Else
        MsgBox "tela export não abriu", vbCritical, "Tela Export"
        End

    End If
    
    'espera tela salvar como que neste caso chama 'Choose Export File'
    GlobalIDTelaSalvarComo = 0
    lcount = 0
    Do While GlobalIDTelaSalvarComo = 0 And lcount < 20000
        lcount = lcount + 1
        GlobalIDTelaSalvarComo = FindWindow("#32770", "Choose Export File")
        DoEvents
        Sleep 300
    Loop
    If GlobalIDTelaSalvarComo = 0 Then MsgBox lcount
    If GlobalIDTelaSalvarComo > 0 Then
        lcount = 0
        hNomedoArquivo = 0
        Do While hNomedoArquivo = 0 Or lcount > 10
            hDUIView = FindWindowEx(GlobalIDTelaSalvarComo, 0, "DUIViewWndClassName", "")
            'encontra DirectUIHWND
            hDirectUI = FindWindowEx(hDUIView, 0, "DirectUIHWND", "")
            'encontra FloatNotifySink
            hFloatNotify = FindWindowEx(hDirectUI, 0, "FloatNotifySink", "")
            'encontra ComboBox
            hComboBox = FindWindowEx(hFloatNotify, 0, "ComboBox", "")
            'encontra caixa texto de nome do arquivo
            hNomedoArquivo = FindWindowEx(hComboBox, 0, "Edit", "")
    
            lcount = lcount + 1
            Sleep 300
            DoEvents
        Loop
        If hNomedoArquivo = 0 Then MsgBox lcount
        SendMessage2 hNomedoArquivo, WM_SETTEXT, 0, GlobalPastadeTrabalho & "\" & GlobalNomedoRelatorio & ".rtf" & Chr$(0)
        DoEvents
        Sleep 300
        'verificar na APS E DIMINUIR O TEMPO
    
        hBotãoSalvar = FindWindowEx(GlobalIDTelaSalvarComo, 0, "Button", "Sa&lvar")
        If hBotãoSalvar = 0 Then Stop
        SendMessage hBotãoSalvar, BM_CLICK, 0, 0
        DoEvents
        Sleep 1000
        SendMessage GlobalIDTelaSalvarComo, WM_CLOSE, 0, 0
    Else
        MsgBox "Tela 'Salvar como' não abriu. Reinicialize o aplicativo."
        End
    End If

    
    
 

    
    'espera o arquivo aparecer
    While Dir(GlobalPastadeTrabalho & "\" & GlobalNomedoRelatorio & ".rtf") = ""
        Sleep 100
        DoEvents
    Wend
    While FileLen(GlobalPastadeTrabalho & "\" & GlobalNomedoRelatorio & ".rtf") = 0
        Sleep 100
        DoEvents
    Wend
    rtbRequerimentos.LoadFile GlobalPastadeTrabalho & "\" & GlobalNomedoRelatorio & ".rtf"
    Sleep 300
    decodeRequerimentos rtbRequerimentos.Text
    'fecha as telas
    SendMessage GlobalIDTelaRequerimentosCrystalReport, WM_CLOSE, 0, 0
    SendMessage GlobalIDTelaImprimirAgendamento, WM_CLOSE, 0, 0


    
End Function
Function SeTelaInternaAtiva(NomedaTela As String) As Long
    Dim lngHWnd As Long
    Dim lngHWnd2 As Long
    Dim titletmp As String
    Dim nret As Long
    Dim size As RECT
    lngHWnd = FindWindowEx(GlobalIDControleOperacional, 0, vbNullString, vbNullString)
    Do Until lngHWnd = 0
        If lngHWnd > 0 Then
            lngHWnd2 = FindWindowEx(lngHWnd, 0, vbNullString, vbNullString)
            Do Until lngHWnd2 = 0
                If lngHWnd2 > 0 Then
                    titletmp = Space(256)
                    nret = GetWindowText(lngHWnd2, titletmp, Len(titletmp))
                    If InStr(1, titletmp, NomedaTela) > 0 Then
                        SeTelaInternaAtiva = lngHWnd2
                        Exit Function
                    End If
                End If
                lngHWnd2 = FindWindowEx(lngHWnd2, lngHWnd, vbNullString, vbNullString)
            Loop
        End If
        lngHWnd = FindWindowEx(GlobalIDControleOperacional, lngHWnd, vbNullString, vbNullString)
    Loop
    SeTelaInternaAtiva = 0
End Function
    Sub MontaListadeRequerimentos(memotexto As String)
        Dim pos As Long
        Dim linha As String
        Dim conta As Long
        Dim indice As Long
        Dim GlobalRequerimentosProv(1000) As Requerimento
        indice = 0
        conta = 0
        memotexto = memotexto & Chr(13) & Chr(10)
        pos = InStr(1, memotexto, Chr(13))
        While pos > 0
            linha = Mid(memotexto, 1, pos - 1)
            memotexto = Mid(memotexto, pos + 2)
            If Len(linha) > 5 Then
                If IsNumeric(linha) Then
                    While Asc(Left$(linha, 1)) < 48 Or Asc(Left$(linha, 1)) > 57
                        linha = Mid(linha, 2)
                    Wend
                    While Asc(Right$(linha, 1)) < 48 Or Asc(Right$(linha, 1)) > 57
                        linha = Mid(linha, 1, Len(linha) - 1)
                    Wend

                    GlobalRequerimentosProv(conta).Número = linha

                Else
                    conta = conta + 1
                    While Asc(Left$(linha, 1)) < 65 Or Asc(Left$(linha, 1)) > 122
                        linha = Mid(linha, 2)
                    Wend
                    While Asc(Right$(linha, 1)) < 65 Or Asc(Right$(linha, 1)) > 122
                        linha = Mid(linha, 1, Len(linha) - 1)
                    Wend

                    GlobalRequerimentosProv(conta).Segurado = linha

                End If
                
            End If
            pos = InStr(2, memotexto, Chr(13))
        Wend
        GlobalQuantidadedeRequerimentos = conta
        lstClassificar.Clear
        For conta = 1 To GlobalQuantidadedeRequerimentos
            lstClassificar.AddItem GlobalRequerimentosProv(conta).Segurado
            lstClassificar.ItemData(lstClassificar.NewIndex) = conta
        Next conta
        lstMostrarRequerimentos.Clear
        lstMostrarRequerimentos.AddItem "Seq. Requerimento Segurado"
        For conta = 0 To lstClassificar.ListCount - 1
            lstMostrarRequerimentos.Visible = True
            indice = indice + 1
            lstMostrarRequerimentos.AddItem Format(indice, "000") & "  " & GlobalRequerimentosProv(lstClassificar.ItemData(conta)).Número & "      " & lstClassificar.List(conta)
            GlobalRequerimentos(conta + 1).Número = GlobalRequerimentosProv(lstClassificar.ItemData(conta)).Número
            GlobalRequerimentos(conta + 1).Segurado = GlobalRequerimentosProv(lstClassificar.ItemData(conta)).Segurado
        Next conta

        If GlobalQuantidadedeRequerimentos > 0 Then
            'lblNit.Visible = True
        Else
            MsgBox "Não foi encontrado nenhum requerimento para ser impresso."
            tmVeriricaSeControleEstaAberto.Enabled = True
            'tmVerificaseMenuRequerimentosfoiAcionado.Enabled = True
            Exit Sub

        End If
        Me.Width = 8000
        Me.Height = 6000
        SetForegroundWindow (Me.hwnd)

    End Sub
Private Function ObtemTelaSecundariaporTituloLarguraeAltura(ByVal sCaption As String, sLargura As Long, sAltura As Long) As Long
    Dim lhWndP As Long
    Dim sStr As String
    Dim size As RECT
    
    ObtemTelaSecundariaporTituloLarguraeAltura = 0
    lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
    Do While lhWndP <> 0
        sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
        GetWindowText lhWndP, sStr, Len(sStr)
        sStr = Left$(sStr, Len(sStr) - 1)
        res = GetWindowRect(lhWndP, size)
        If InStr(1, sStr, sCaption) > 0 Then
            If size.Right - size.Left = sLargura And size.Bottom - size.Top = sAltura Then
                ObtemTelaSecundariaporTituloLarguraeAltura = lhWndP
                Exit Function
            End If
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop
End Function
    
Private Function ObtemTelaPrincipalporTitulo(ByVal sCaption As String) As Long
    Dim lhWndP As Long
    Dim sStr As String
    ObtemTelaPrincipalporTitulo = False
    lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
    Do While lhWndP <> 0
        sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
        GetWindowText lhWndP, sStr, Len(sStr)
        sStr = Left$(sStr, Len(sStr) - 1)
        
        If InStr(1, sStr, sCaption) > 0 Then
            ObtemTelaPrincipalporTitulo = lhWndP
            Exit Function
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop
    ObtemTelaPrincipalporTitulo = 0
End Function
    Sub MouseClique(posx As Long, posy As Long)
        Dim size As RECT
        Dim IDTelaAtiva As Long
        Dim pt As POINTAPI
        GetCursorPos pt
        IDTelaAtiva = GetForegroundWindow
        res = GetWindowRect(IDTelaAtiva, size)
        SetCursorPos size.Left + posx, size.Top + posy
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        SetCursorPos pt.X, pt.Y
    End Sub
    
    Sub ListaItensdeMenu()
        Dim lngHWnd As Long
        Dim lngHWnd2 As Long
        Dim lngHWnd3 As Long
        Dim lngHWnd4 As Long
        Dim titletmp As String
        Dim nret As Long
        Dim size As RECT
        Dim res As String
        Dim strClass As String
        lstClassificar.Clear
        lngHWnd = FindWindowEx(GlobalIDControleOperacional, 0, vbNullString, vbNullString)
     
        Do Until lngHWnd = 0
            If lngHWnd > 0 Then
                    lngHWnd2 = FindWindowEx(lngHWnd, 0, vbNullString, vbNullString)
                    Do Until lngHWnd2 = 0
                        
    
                        If lngHWnd2 > 0 Then
                            titletmp = Space(256)
                            nret = GetWindowText(lngHWnd2, titletmp, Len(titletmp))
                            lstClassificar.AddItem titletmp
    
                        End If
                        lngHWnd2 = FindWindowEx(lngHWnd2, lngHWnd, vbNullString, vbNullString)
    
                    Loop
    
            End If
                
                'res = GetWindowRect(lngHWnd, size)
            'MsgBox size.Left & " " & size.Right & "  " & size.Top & " " & size.Bottom
        'res = SetWindowPos(lngHWnd, 0, 200, 200, 800, 460, 0)
        'End If
        lngHWnd = FindWindowEx(GlobalIDControleOperacional, lngHWnd, vbNullString, vbNullString)
        Loop
    End Sub
    
 Sub Get_User_Name()
        On Error Resume Next 'voltar
        Dim lpBuff As String * 25
        Dim ret As Long
        ' Get the user name minus any trailing spaces found in the name.
        ret = GetUserName(lpBuff, 25)
        GlobalUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
        'GlobalPastadeTrabalho = "c:\Users\" & GlobalUserName & "\AppData\Local"
        GlobalPastadeTrabalho = "c:\Users\" & GlobalUserName & "\Desktop"

 
    End Sub
    
Private Sub ClickMenu(lAplicativo As Long, lMenu As Long, lItem As Long)
    Dim lSubMenu  As Long
    Dim lMenuItem As Long
    Dim lIDMenu As Long

    'This is a bit more interesting

    lIDMenu = GetMenu(lAplicativo)
    lSubMenu = GetSubMenu(lIDMenu, lMenu)
    lMenuItem = GetMenuItemID(lSubMenu, lItem)

    Call PostMessage(lAplicativo, WM_COMMAND, lMenuItem, 0)
    'sendmessage would hang app until file is selected in open form but
    'postmessage is asynchronous which is better in this case
End Sub

Private Sub GetMenuInfo(hMenu As Long, spaces As Integer, txt As String)
Dim num As Integer
Dim i As Integer
Dim Length As Long
Dim sub_hmenu As Long
Dim sub_name As String
    
    num = GetMenuItemCount(hMenu)
    For i = 0 To num - 1
        ' Save this menu's info.
        sub_hmenu = GetSubMenu(hMenu, i)
        sub_name = Space$(256)
        Length = GetMenuString(hMenu, i, sub_name, Len(sub_name), MF_BYPOSITION)
        sub_name = Left$(sub_name, Length)

        txt = txt & Space$(spaces) & sub_name & vbCrLf
        
        ' Get its child menu's names.
        GetMenuInfo sub_hmenu, spaces + 4, txt
    Next i
End Sub
Private Sub ImprimeSegundaVia(Requerimento As String)
    Dim resposta, respostarequerimento As String
    Dim res As String
    Dim hBotãoOKPesquisaAvançada As Long
    Dim hCampoRequerimentoPesquisaAvançada As Long
    Dim MarcaçãoImpressa As String
    Dim contavezes As Long
    Dim size As RECT
    Dim lParam As Long
    Dim EncontraValorNIT As String
    GlobalRelaçaodeRequerimentosImpressos = ""
    GlobalHoradeInicio = Time
    For GlobalIDRequerimento = 1 To GlobalQuantidadedeRequerimentos
        If GlobalRequerimentos(GlobalIDRequerimento).Número = Trim(Requerimento) Then
            GlobalIDTelaPesquisaAvançada = 0
            GlobalMehwndMSG = Me.hwnd
            GlobalNumerodoRequerimentoMSG = GlobalRequerimentos(GlobalIDRequerimento).Número
            GlobalNomedoSeguradoMSG = GlobalRequerimentos(GlobalIDRequerimento).Segurado
            atualizaprogresso
            
            resposta = ""
            resposta = abrePesquisaAvançada
            If GlobalIDTelaPesquisaAvançada <> 0 Then
                respostarequerimento = PreencheCampoRequerimento(GlobalIDTelaPesquisaAvançada)
                If respostarequerimento = "Sucesso" Then
                    'Text1.Text = msgboxRequerimento
                    'verifica se tela 'Carteira de Benefícios' apareceu
                    'se nunca vier basta clicar no mouse esquerdo para sair do loop
                    While GetAsyncKeyState(VK_LBUTTON) = True
                    Wend

                    While GlobalIDTelaCarteiradeBeneficios = 0 And GetAsyncKeyState(VK_LBUTTON) = False
                        GlobalIDTelaCarteiradeBeneficios = ObtemTelaPrincipalporTitulo("Carteira de Benefícios")
                    Wend
                    While GlobalIDTelaCarteiradeBeneficios <> 0
                        GlobalIDTelaCarteiradeBeneficios = ObtemTelaPrincipalporTitulo("Carteira de Benefícios")
                        Sleep 100
                        DoEvents
                    Wend
                    
                    If chkImpressãoautomática.Value = False Then
                        Sleep 2000
                        resposta = msgboxInformaçoesdoRequerimento
                    Else
                        Sleep 4000
                        resposta = "Completo"
                
                    End If
                    
                    If resposta = "Vazio" Then
                        msgboxVazio
                        'MsgBox "O requerimento '" & GlobalNumerodoRequerimentoMSG & "-" & GlobalNomedoSeguradoMSG & "' deve ser de outra APS. Não será possibel extrair o NIT.", vbCritical, "Requerimento de Outra APS"
                        GlobalRequerimentos(GlobalIDRequerimento).NIT = "0.000.000.000-0"
                        atualizalista
                        GoTo PROXIMOREQUERIMENTO
                    End If
                    If resposta = "Completo" Then
                        
                        'fecha tela Pesquisa Avançado
                        'clica em Serviço
                        MouseCliqueAbsoluto 855, 300
                        DoEvents
                        Sleep 300
                        'clica em Nit secundario
                        MouseCliqueAbsoluto 927, 668
                        DoEvents
                        EncontraValorNIT = ExtraiNIT
                        If EncontraValorNIT <> GlobalUltimoNitInformado Then
                            GlobalUltimoNitInformado = EncontraValorNIT
                            GlobalRequerimentos(GlobalIDRequerimento).NIT = EncontraValorNIT
                            atualizalista
                            If EncontraValorNIT <> "0.000.000.000-0" Then
                                
                                MarcaçãoImpressa = AbreSegundaViaMarcaçãodeExame(EncontraValorNIT)
                                If MarcaçãoImpressa <> "" Then MsgBox MarcaçãoImpressa, vbCritical, "2ª Via de Marcação de Exame"
                            Else
                                MsgBox "NIT inválido"
                            End If
                        Else
                            MsgBox "Foi encontrado o mesmo NIT do requerimento anterior. Provavelmente você não aguardou as informações do novo requerimento.", vbCritical, "NIT Repetido"
                        End If
                    End If
    
                    'Text1.Text = resposta
                    'Text1.Text = msgboxMarcaçãodeExame
                End If
            Else
                If resposta = "A tela 'Pesquisa Avançada' não abriu." Then
                    MsgBox resposta, vbCritical, "Pesquisa Avançada"
                    Exit Sub
                End If
            End If
            
PROXIMOREQUERIMENTO:
            'sem esta rotina a verificação de Pesquisa Avançada volta positiva
            contavezes = ObtemTelaPrincipalporTitulo("Pesquisa Avançada")
            While contavezes <> 0
                SendMessage contavezes, WM_CLOSE, 0, 0
                contavezes = ObtemTelaPrincipalporTitulo("Pesquisa Avançada")
                DoEvents
                Sleep 100
            Wend
        End If
    Next GlobalIDRequerimento
End Sub






Private Sub RotinaImprimeSegundaVia()
    Dim resposta, respostarequerimento As String
    Dim res As String
    Dim hBotãoOKPesquisaAvançada As Long
    Dim hCampoRequerimentoPesquisaAvançada As Long
    Dim MarcaçãoImpressa As String
    Dim contavezes As Long
    Dim size As RECT
    Dim lParam As Long
    Dim EncontraValorNIT As String
    GlobalRelaçaodeRequerimentosImpressos = ""
    GlobalHoradeInicio = Time
    For GlobalIDRequerimento = 1 To GlobalQuantidadedeRequerimentos
        GlobalIDTelaPesquisaAvançada = 0
        GlobalMehwndMSG = Me.hwnd
        GlobalNumerodoRequerimentoMSG = GlobalRequerimentos(GlobalIDRequerimento).Número
        GlobalNomedoSeguradoMSG = GlobalRequerimentos(GlobalIDRequerimento).Segurado
        atualizaprogresso
        
        resposta = ""
        resposta = abrePesquisaAvançada
        If GlobalIDTelaPesquisaAvançada <> 0 Then
            respostarequerimento = PreencheCampoRequerimento(GlobalIDTelaPesquisaAvançada)
            If respostarequerimento = "Sucesso" Then
                'Text1.Text = msgboxRequerimento
                'verifica se tela 'Carteira de Benefícios' apareceu
                'se nunca vier basta clicar no mouse esquerdo para sair do loop
                While GetAsyncKeyState(VK_LBUTTON) = True
                Wend
                While GlobalIDTelaCarteiradeBeneficios = 0 And GetAsyncKeyState(VK_LBUTTON) = False
                    GlobalIDTelaCarteiradeBeneficios = ObtemTelaPrincipalporTitulo("Carteira de Benefícios")
                Wend
                While GlobalIDTelaCarteiradeBeneficios <> 0
                    GlobalIDTelaCarteiradeBeneficios = ObtemTelaPrincipalporTitulo("Carteira de Benefícios")
                    Sleep 100
                    DoEvents
                Wend
                
                If chkImpressãoautomática.Value = False Then
                    Sleep 2000
                    resposta = msgboxInformaçoesdoRequerimento
                Else
                    Sleep 4000
                    resposta = "Completo"
            
                End If
                
                If resposta = "Vazio" Then
                    msgboxVazio
                    'MsgBox "O requerimento '" & GlobalNumerodoRequerimentoMSG & "-" & GlobalNomedoSeguradoMSG & "' deve ser de outra APS. Não será possibel extrair o NIT.", vbCritical, "Requerimento de Outra APS"
                    GlobalRequerimentos(GlobalIDRequerimento).NIT = "0.000.000.000-0"
                    atualizalista
                    GoTo PROXIMOREQUERIMENTO
                End If
                If resposta = "Completo" Then
                    
                    'fecha tela Pesquisa Avançado
                    'clica em Serviço
                    MouseCliqueAbsoluto 855, 300
                    DoEvents
                    Sleep 300
                    'clica em Nit secundario
                    MouseCliqueAbsoluto 927, 668
                    DoEvents
                    EncontraValorNIT = ExtraiNIT
                    If EncontraValorNIT <> GlobalUltimoNitInformado Then
                        GlobalUltimoNitInformado = EncontraValorNIT
                        GlobalRequerimentos(GlobalIDRequerimento).NIT = EncontraValorNIT
                        atualizalista
                        If EncontraValorNIT <> "0.000.000.000-0" Then
                            
                            MarcaçãoImpressa = AbreSegundaViaMarcaçãodeExame(EncontraValorNIT)
                            If MarcaçãoImpressa <> "" Then MsgBox MarcaçãoImpressa, vbCritical, "2ª Via de Marcação de Exame"
                        Else
                            MsgBox "NIT inválido"
                        End If
                    Else
                        MsgBox "Foi encontrado o mesmo NIT do requerimento anterior. Provavelmente você não aguardou as informações do novo requerimento.", vbCritical, "NIT Repetido"
                    End If
                End If

                'Text1.Text = resposta
                'Text1.Text = msgboxMarcaçãodeExame
            End If
        Else
            If resposta = "A tela 'Pesquisa Avançada' não abriu." Then
                MsgBox resposta, vbCritical, "Pesquisa Avançada"
                Exit Sub
            End If
        End If
        
PROXIMOREQUERIMENTO:
        'sem esta rotina a verificação de Pesquisa Avançada volta positiva
        contavezes = ObtemTelaPrincipalporTitulo("Pesquisa Avançada")
        While contavezes <> 0
            SendMessage contavezes, WM_CLOSE, 0, 0
            contavezes = ObtemTelaPrincipalporTitulo("Pesquisa Avançada")
            DoEvents
            Sleep 100
        Wend

    Next GlobalIDRequerimento
End Sub



Private Sub cmdImprimir_Click()
    Dim resposta, respostarequerimento As String
    Dim res As String
    Dim hBotãoOKPesquisaAvançada As Long
    Dim hCampoRequerimentoPesquisaAvançada As Long
    Dim MarcaçãoImpressa As String
    Dim contavezes As Long
    Dim size As RECT
    Dim lParam As Long
    Dim EncontraValorNIT As String
    modoImprime = "Todos"
    GlobalRelaçaodeRequerimentosImpressos = ""
    GlobalHoradeInicio = Time
    For GlobalIDRequerimento = 1 To GlobalQuantidadedeRequerimentos
        GlobalIDTelaPesquisaAvançada = 0
        GlobalMehwndMSG = Me.hwnd
        GlobalNumerodoRequerimentoMSG = GlobalRequerimentos(GlobalIDRequerimento).Número
        GlobalNomedoSeguradoMSG = GlobalRequerimentos(GlobalIDRequerimento).Segurado
        atualizaprogresso
        
        resposta = ""
        resposta = abrePesquisaAvançada
        If GlobalIDTelaPesquisaAvançada <> 0 Then
            respostarequerimento = PreencheCampoRequerimento(GlobalIDTelaPesquisaAvançada)
            If respostarequerimento = "Sucesso" Then
                'Text1.Text = msgboxRequerimento
                'verifica se tela 'Carteira de Benefícios' apareceu
                'se nunca vier basta clicar no mouse esquerdo para sair do loop
                While GetAsyncKeyState(VK_LBUTTON) = True
                Wend

                While GlobalIDTelaCarteiradeBeneficios = 0 And GetAsyncKeyState(VK_LBUTTON) = False
                    GlobalIDTelaCarteiradeBeneficios = ObtemTelaPrincipalporTitulo("Carteira de Benefícios")
                Wend
                While GlobalIDTelaCarteiradeBeneficios <> 0
                    GlobalIDTelaCarteiradeBeneficios = ObtemTelaPrincipalporTitulo("Carteira de Benefícios")
                    Sleep 100
                    DoEvents
                Wend
                
                If chkImpressãoautomática.Value = False Then
                    Sleep 2000
                    resposta = msgboxInformaçoesdoRequerimento
                Else
                    Sleep 4000
                    resposta = "Completo"
            
                End If
                
                If resposta = "Vazio" Then
                    msgboxVazio
                    'MsgBox "O requerimento '" & GlobalNumerodoRequerimentoMSG & "-" & GlobalNomedoSeguradoMSG & "' deve ser de outra APS. Não será possibel extrair o NIT.", vbCritical, "Requerimento de Outra APS"
                    GlobalRequerimentos(GlobalIDRequerimento).NIT = "0.000.000.000-0"
                    atualizalista
                    GoTo PROXIMOREQUERIMENTO
                End If
                If resposta = "Completo" Then
                    
                    'fecha tela Pesquisa Avançado
                    'clica em Serviço
                    MouseCliqueAbsoluto 855, 300
                    DoEvents
                    Sleep 300
                    'clica em Nit secundario
                    MouseCliqueAbsoluto 927, 668
                    DoEvents
                    EncontraValorNIT = ExtraiNIT
                    If EncontraValorNIT <> GlobalUltimoNitInformado Then
                        GlobalUltimoNitInformado = EncontraValorNIT
                        GlobalRequerimentos(GlobalIDRequerimento).NIT = EncontraValorNIT
                        atualizalista
                        If EncontraValorNIT <> "0.000.000.000-0" Then
                            
                            MarcaçãoImpressa = AbreSegundaViaMarcaçãodeExame(EncontraValorNIT)
                            If MarcaçãoImpressa <> "" Then MsgBox MarcaçãoImpressa, vbCritical, "2ª Via de Marcação de Exame"
                        Else
                            MsgBox "NIT inválido"
                        End If
                    Else
                        MsgBox "Foi encontrado o mesmo NIT do requerimento anterior. Provavelmente você não aguardou as informações do novo requerimento.", vbCritical, "NIT Repetido"
                    End If
                End If

                'Text1.Text = resposta
                'Text1.Text = msgboxMarcaçãodeExame
            End If
        Else
            If resposta = "A tela 'Pesquisa Avançada' não abriu." Then
                MsgBox resposta, vbCritical, "Pesquisa Avançada"
                Exit Sub
            End If
        End If
        
PROXIMOREQUERIMENTO:
        'sem esta rotina a verificação de Pesquisa Avançada volta positiva
        contavezes = ObtemTelaPrincipalporTitulo("Pesquisa Avançada")
        While contavezes <> 0
            SendMessage contavezes, WM_CLOSE, 0, 0
            contavezes = ObtemTelaPrincipalporTitulo("Pesquisa Avançada")
            DoEvents
            Sleep 100
        Wend

    Next GlobalIDRequerimento

End Sub

Private Sub cmdImprimir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdImprimir.BorderStyle = 1
End Sub

Private Sub cmdImprimir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdImprimir.BorderStyle = 0
End Sub

Private Sub cmdImprimir1_Click()

End Sub

Private Sub Form_Activate()
    tmVeriricaSeControleEstaAberto.Enabled = True
End Sub

Private Sub Form_Load()
    Dim res As String
    Get_User_Name
    res = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
    Me.Top = 4300
    Me.Left = 1200
    Me.Width = 4000
    Me.Height = 133 * 15
    GlobalTítulodaTelaAtiva = ""
    GlobalMenuAtualizado = False
    GlobalIDControleOperacional = 0
    GlobalModoImprimeRequerimentos = False
    GlobalEscalaX = 256 / Screen.Width
    GlobalEscalaX = GlobalEscalaX * 256
    GlobalEscalay = 256 / Screen.Height
    GlobalEscalay = GlobalEscalay * 256
    pctProgresso.Width = 0
    pctFundoProgresso.Visible = False
    
      
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    pctFundo.Top = 0
    pctFundo.Left = 0
    pctFundoProgresso.Left = 240
    pctFundoProgresso.Width = Me.Width - 580
    pctFundo.Width = Me.Width
    lblRequerimentodoSABI.Width = pctFundo.Width
    pctFundo.Height = Me.Height
    lblRequerimentodoSABI.Top = 80
    rtbRequerimentos.Width = Me.Width - 600
    pctFundoProgresso.Top = lblRequerimentodoSABI.Top + lblRequerimentodoSABI.Height + 40
    lstMostrarRequerimentos.Top = pctFundoProgresso.Top + pctFundoProgresso.Height + 40
    lstMostrarRequerimentos.Left = 240
    lstMostrarRequerimentos.Width = Me.Width - 580
    lstMostrarRequerimentos.Height = Me.Height - lstMostrarRequerimentos.Top - 600
    
    chkImpressãoautomática.Top = pctFundo.Height - cmdImprimir.Height - 240
    chkImpressãoautomática.Left = 240

    
    cmdImprimir.Top = pctFundo.Height - cmdImprimir.Height - 240
    cmdImprimir.Left = 3000
    imgFecha.Top = cmdImprimir.Top - 120
    imgFecha.Left = pctFundo.Width - imgFecha.Width - 240
    lblversão.Top = pctFundo.Height - 300
    lblversão.Left = cmdImprimir.Left + 3100

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim res As String
    res = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3)

    MsgBox "Encerrando o aplicativo 'Requerimentos do SABI'.", vbCritical, "Requerimentos do SABI"
    End
End Sub

Private Sub imgFecha_Click()
    Unload Me
End Sub

Private Sub lblMarcaTodos_Click()
    Dim conta As Long
    For conta = 1 To lstClassificar.ListCount
        lstMostrarRequerimentos.Selected(conta) = True
    Next conta

End Sub

Private Sub lblImprimir_Click()
    Dim conta As Long
    Dim size As RECT
    SetForegroundWindow (GlobalIDControleOperacional)
    For conta = 1 To GlobalQuantidadedeRequerimentos
        If GlobalRequerimentos(conta).NIT <> "" Then
            'Abre tela Segunda Via de Marcação de Exame
            
            SetForegroundWindow (GlobalIDControleOperacional)
 
            ClickMenu GlobalIDControleOperacional, 4, 7
            Sleep 1000
            SendKeys "{TAB}"
            Sleep 100
            SendKeys "{TAB}"
            Sleep 100
            SendKeys GlobalRequerimentos(conta).NIT
            Sleep 100
            SendKeys Left$(GlobalRequerimentos(conta).NIT, 1)
            Sleep 300
            MsgBox "imprime"
            'GlobalIDTelaSegundaVia = achaTelaInternaAtiva("Segunda Via de Marcação de Exame")
            'res = GetWindowRect(GlobalIDTelaSegundaVia, size)
            'While size.Right - size.Left < 0
                'sleep  100
                'DoEvents
                'GlobalIDTelaSegundaVia = achaTelaInternaAtiva("Segunda Via de Marcação de Exame")
                'res = GetWindowRect(GlobalIDTelaSegundaVia, size)
            'Wend

            
            
        End If
    Next conta

End Sub

Private Sub lblImprimir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblImprimir.BorderStyle = 1
End Sub

Private Sub lblImprimir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblImprimir.BorderStyle = 0
End Sub

Private Sub pesquisalNit()

    On Error Resume Next
'qual o nome da tela CarteiradeBeneficios?Consulta Requerimento/Benefício
'qual o menu que a abre? 2 0
'qual a largura e altura de GloblaIDTelaRelatorioEmitidocomSucesso
    Dim conta As Long
    Dim strClass As String
    Dim nret As Long
    Dim lhWndP As Long
    Dim arquivo As String
    Dim memoerro As String
    Dim hDC As Long
    Dim memo As String
    Dim size As RECT
    Dim hBotaoSairPesquisa As Long
    Dim hBotaoSairNIT As Long
    Dim hValorRequerimento As Long
    Dim hValorNIT As Long
    Dim BotãoTexto As String
    Dim ValorNit As String
    Dim hPrimeira As Long
    Dim ClassedoControle As String
    Dim classlength As Long
    Dim contaloop As Long
    Dim hThunderRT6FrameIMPRIME As Long
    Dim hImMaskWndCIassIMPRIME As Long
    memoerro = ""

    GlobalHoradeInicio = Time
    hDC = GetWindowDC(0)
    memoerro = ""
    'lstClassificar.AddItem GlobalRequerimentos(conta).Número & " - " & GlobalRequerimentos(conta).Segurado
    SetForegroundWindow (GlobalIDControleOperacional)
    DoEvents
    Sleep 300
    'verifica se tela 'Carteira de Benefícios' esta aberta
    GlobalIDTelaCarteiradeBeneficios = SeTelaInternaAtiva("Consulta Requerimento/Benefício")
    If GlobalIDTelaCarteiradeBeneficios = 0 Then
        ClickMenu GlobalIDControleOperacional, 2, 0
        While GlobalIDTelaCarteiradeBeneficios = 0
            Sleep 100
            DoEvents
            GlobalIDTelaCarteiradeBeneficios = SeTelaInternaAtiva("Consulta Requerimento/Benefício")
        Wend
    End If
    'Call SendMessage(GlobalIDTelaCarteiradeBeneficios, WM_NCPAINT, 0&, 0&)
    Sleep 1000
    memo = "Consulta Requerimento/Benefício ativa"
    'TextOut hDC, Screen.Width / 30, (Screen.Height - 600) / 15 + 20, memo, Len(memo)
    DoEvents
    For GlobalIDRequerimento = 1 To GlobalQuantidadedeRequerimentos
        pctFundoProgresso.Visible = True
        memo = CDate(Time - GlobalHoradeInicio) & " > " & Format(GlobalIDRequerimento, "000") & "/" & Format(GlobalQuantidadedeRequerimentos, "000")
        lblProgresso.Caption = memo & " " & GlobalRequerimentos(GlobalIDRequerimento).Segurado
        lblProgresso2.Caption = memo & " " & GlobalRequerimentos(GlobalIDRequerimento).Segurado
        pctProgresso.Width = pctFundoProgresso.Width * (GlobalIDRequerimento / GlobalQuantidadedeRequerimentos)
        DoEvents
        'res = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)

        GlobalNomedoRelatorio = "nit" & GlobalRequerimentos(GlobalIDRequerimento).Número
        'sleep  600
        'Abre tela Pesquisa Avançada
        AbrirTelaPesquisaAvançada GlobalRequerimentos(GlobalIDRequerimento).Número
        
        

            'SetForegroundWindow (GlobalIDTelaCarteiradeBeneficios)
            Sleep 2000
            MsgBox "Clique OK quando aparecer '" & GlobalRequerimentos(GlobalIDRequerimento).Segurado & "'"
            'sleep  2000
            'clica em seta 'imprimir'
            DoEvents
            'sleep  300
            
            'abre menu Serviços
            'SetCursorPos 855, 300
            DoEvents
            Sleep 1000
            MouseCliqueAbsoluto 855, 300
            DoEvents
            'SetCursorPos 927, 668
            DoEvents
            'MouseClique 855, 300
            Sleep 1000
            
            'clica em Nit secundario
            MouseCliqueAbsoluto 927, 668
            DoEvents
            Sleep 300
            
            
            'verifica se tela Informações de NIT(s) Secundário(s)' esta aberta
            contaloop = 0
            GlobalIDTelaNITSecundario = 0
            While GlobalIDTelaNITSecundario = 0 And contaloop < 50
                GlobalIDTelaNITSecundario = ObtemTelaPrincipalporTitulo("Informações de NIT(s) Secundário(s)")
                Sleep 300
                DoEvents
                GlobalIDTelaConsultaSemCriterio = ObtemTelaPrincipalporTitulo("AVISO IMPORTANTE")
                If GlobalIDTelaConsultaSemCriterio <> 0 Then
                    MsgBox "CONSULTA SEM CRITERIO"
                    End
                End If
                contaloop = contaloop + 1
                'NO dia 08/07/2015 APS/DIVINÓPOLIS aparece entre os requerimento
                '15:20        EDGAR RIBEIRO  DE OLIVEIRA  MORAES           N       166101916
                'APS -  BELO HORIZONTE-OESTE
                'este loop fica rodando sem fim
            Wend
            If contaloop = 50 Then
                ValorNit = "0.000.000.000-0"
                GoTo novovalordenit
            End If
                
            hValorNIT = 0
            While hValorNIT = 0 And contaloop < 50
                hPrimeira = FindWindowEx(GlobalIDTelaNITSecundario, 0, vbNullString, vbNullString)
                Do While hPrimeira <> 0 And contaloop < 50
                    ClassedoControle = Space(128)
                    classlength = GetClassName(hPrimeira, ClassedoControle, 128)
                    ClassedoControle = Left(ClassedoControle, classlength)
                    If ClassedoControle = "ThunderRT6CommandButton" Then hBotaoSairNIT = hPrimeira
                    If ClassedoControle = "ImMaskWndClass" Then hValorNIT = hPrimeira
                    hPrimeira = FindWindowEx(GlobalIDTelaNITSecundario, hPrimeira, vbNullString, vbNullString)
                    'sleep  300
                    DoEvents
                    contaloop = contaloop + 1
                Loop
            Wend
            Sleep 300
            ValorNit = WindowTextGet(hValorNIT)
            Sleep 300
        'Else
            'ValorNIT = memoerro
        'End If
        SendMessage hBotaoSairNIT, BM_CLICK, 0, 0
        Sleep 600
        SendMessage GlobalIDTelaNITSecundario, WM_CLOSE, 0, 0
        If ValorNit <> "0.000.000.00-0" Then
        
        
        'abre tela de impresssao
        While InStr(1, ValorNit, ".")
            ValorNit = Mid(ValorNit, 1, InStr(1, ValorNit, ".") - 1) & Mid(ValorNit, InStr(1, ValorNit, ".") + 1)
        Wend
        'ValorNIT = Mid(ValorNIT, 1, Len(ValorNIT) - 2)
        ClickMenu GlobalIDControleOperacional, 4, 7
        While GetAsyncKeyState(VK_LBUTTON) = True
        Wend '

        While GetAsyncKeyState(VK_LBUTTON) = False
        Wend '
        While GetAsyncKeyState(VK_LBUTTON) = True
        Wend '
        'MsgBox "clique ok QUANDO a tela de impressao aparecer"
        contaloop = 0
        GlobalIDTelaSegundaVia = 0
        While GlobalIDTelaSegundaVia = 0 And contaloop < 100
            GlobalIDTelaSegundaVia = SeTelaInternaAtiva("Segunda Via de Marcação de Exame")
            Sleep 300
            DoEvents
            contaloop = contaloop + 1
        Wend
        DoEvents
        Sleep 300
        res = SetWindowPos(GlobalIDTelaSegundaVia, 0, 0, 0, 800, 200, 0)
        DoEvents
        Sleep 300
        DoEvents
        
        hThunderRT6FrameIMPRIME = FindWindowEx(GlobalIDTelaSegundaVia, 0, "ThunderRT6Frame", "NIT Requerente")
        If hThunderRT6FrameIMPRIME <> 0 Then
            hImMaskWndCIassIMPRIME = FindWindowEx(hThunderRT6FrameIMPRIME, 0, vbNullString, vbNullString)
            strClass = Space(100)
            nret = GetClassName(hImMaskWndCIassIMPRIME, strClass, 100)
            If Mid(strClass, 1, 14) = "ImMaskWndClass" Then
                SendMessage2 hImMaskWndCIassIMPRIME, WM_SETTEXT, 0, ValorNit & Chr$(0)
                DoEvents
            Else
        End If
            End If
            
            
        Clipboard.SetText ValorNit
        MsgBox "Dê um <CONTROL><ALT><V> no campo NIT" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Clique em OK somente quando a marcação de exame estiver impressa", vbApplicationModal, "Segunda Via de Marcação de Exame"
        
        MouseClique 92, 126
        DoEvents
        Sleep 300
        DoEvents
        SendKeys ValorNit
        DoEvents
        SetForegroundWindow (GlobalIDControleOperacional)
        SetCursorPos 92, 126
        MouseClique 92, 126
        SendKeys ValorNit
        
        End If
        
novovalordenit:
        GlobalRequerimentos(GlobalIDRequerimento).NIT = ValorNit
        TextOut hDC, Screen.Width / 30, (Screen.Height - 600) / 15, "   " & ValorNit & "   ", Len(ValorNit) + 6
        atualizalista


        
        
        
    Next GlobalIDRequerimento
    pctProgresso.Width = 0
    pctFundoProgresso.Visible = False
    DoEvents
    Me.Width = 8000
    Me.Height = 6000 + 2000
    res = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
    DoEvents
    memo = "AGENDAMENTOS DO SABI"
    memo = memo & Chr(13) & Chr(10) & "Requerimento" & Chr(9) & "NIT" & Chr(9) & "SEGURADO"
    For conta = 1 To GlobalQuantidadedeRequerimentos
        memo = memo & Chr(13) & Chr(10) & GlobalRequerimentos(conta).Número & Chr(9) & GlobalRequerimentos(conta).NIT & Chr(9) & GlobalRequerimentos(conta).Segurado
    Next conta
    Clipboard.Clear
    Clipboard.SetText memo
    MsgBox "Inicio: " & GlobalHoradeInicio & ", final:" & Time & ", quant.: " & GlobalQuantidadedeRequerimentos & Chr(13) & Chr(10) & "A relação de agendamentos foi salva na área de transferência do computadodor." & Chr(13) & Chr(10) & "Cole no editor de texto."
    End


End Sub



Private Sub lstMostrarRequerimentos_Click()
    Dim memorequerimento As String
    If imprimereq = True Then
        modoImprime = "Único"
        memorequerimento = Mid(lstMostrarRequerimentos, 6)
        memorequerimento = Trim(Mid(memorequerimento, 1, 9))
        ImprimeSegundaVia (memorequerimento)
    End If

End Sub

Private Sub lstMostrarRequerimentos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        imprimereq = True
End Sub


Private Sub lstMostrarRequerimentos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imprimereq = False
End Sub

Private Sub mnImprimir_Click()
    Dim res As String
    Me.Top = -10000
    MsgBox "imprimir"
    Me.Top = 4300

End Sub




Private Sub Timer1_Timer()

End Sub





Private Sub Text1_Change()

End Sub

Private Sub pctMensagem_Click()

End Sub

Private Sub tmImprimirMarcação_Timer()
    Dim conta As Long
    Dim memo As String
    Dim pt As POINTAPI
    Dim hDC As Long
    hDC = GetWindowDC(0)

    GlobalIDTelaImprime = SeTelaInternaAtiva("Segunda Via de Marcação de Exame")
    If GlobalIDTelaImprime > 0 Then
        For conta = 1 To GlobalQuantidadedeRequerimentos
            If GlobalRequerimentos(conta).NIT <> "" Then
                If GlobalRequerimentos(conta).Impresso = False Then
                    GlobalPróximoNITaserimpresso = GlobalRequerimentos(conta).NIT
                    memo = GlobalRequerimentos(conta).NIT
                    TextOut hDC, 0, 0, memo, Len(memo)

                End If
            End If
        Next conta
        GetCursorPos pt
        'DCT/CI: 127 127
        'tecla Imprimir: 536 197
        If pt.X > 120 And pt.X < 140 And pt.Y > 120 And pt.Y > 130 Then
        End If
    End If
End Sub

Private Sub tmMostraPosCursor_Timer()
'desativado
        Dim IDTelaAtiva As Long
        Dim pt As POINTAPI
        Dim size As RECT
        Dim titletmp As String
        Dim nret As Long
        Dim hDC As Long
        Dim memo As String
        hDC = GetWindowDC(0)
        
        IDTelaAtiva = GetForegroundWindow
        res = GetWindowRect(IDTelaAtiva, size)
        titletmp = Space(256)
        nret = GetWindowText(IDTelaAtiva, titletmp, Len(titletmp))
        GetCursorPos pt
        memo = "x: " & pt.X - size.Left & " y: " & pt.Y - size.Top & " larg: " & size.Right - size.Left & " alt: " & size.Bottom - size.Top & " ml: " & GetAsyncKeyState(1) & " tit: " & titletmp
        'TextOut hDC, 0, 0, memo, Len(memo)
        
        'sai se o controle for desligado
        GlobalIDControleOperacional = ObtemTelaPrincipalporTitulo("SABI - Módulo de Controle Operacional")
        res = GetWindowRect(GlobalIDControleOperacional, size)
        If size.Right - size.Left = 0 Then
            MsgBox "Favor abrir o Controle Operacional", vbApplicationModal + vbCritical, "Controle Operacional Fechado"
            End
        End If
End Sub

Private Sub tmRelaçãodeRequerimentos_Timer()
'    On Error Resume Next
    Dim COsize As RECT
    Dim size As RECT
    Dim titletmp As String
    Dim nret As Long
    Dim TelaSize As RECT
    Dim arquivo As String
    Dim memo As String
    On Error Resume Next
    tmRelaçãodeRequerimentos.Enabled = False
    GlobalNomedoRelatorio = "requerimentos"
    
    memo = Dir(GlobalPastadeTrabalho & "\Requerimentos.rtf")
    If memo <> "" Then
        Kill GlobalPastadeTrabalho & "\" & memo
        If Err Then
            MsgBox "Não foi possível apagar a agenda '" & GlobalPastadeTrabalho & "\Requerimentos.rtf'.", vbCritical, "Apagar Agenda Anterior"
            End
        End If
    End If
    'If Dir(GlobalPastadeTrabalho & "\" & GlobalNomedoRelatorio & ".rtf") <> "" Then Kill GlobalPastadeTrabalho & "\" & GlobalNomedoRelatorio & ".rtf"
    'espera a proxima tela (a tela do crystal report não tem nome)
    While GlobalIDTelaImprimirAgendamento = GetForegroundWindow
        Sleep 100
        DoEvents
    Wend

    arquivo = esperaCRYSTALREPORTeExporta
    
    'pesquisalNit
    cmdImprimir.Visible = True
    chkImpressãoautomática.Visible = True
 
End Sub




Private Sub tmVerificaseMenuRequerimentosfoiAcionado_Timer()
    Dim pt As POINTAPI
    Dim COsize As RECT
    If GlobalModoImprimeRequerimentos = False Then
        If GetAsyncKeyState(1) < 0 Then
            GetCursorPos pt
            res = GetWindowRect(GlobalIDControleOperacional, COsize)
            If pt.X - COsize.Left > 494 And pt.X - COsize.Left < 623 And pt.Y - COsize.Top > 30 And pt.Y - COsize.Top < 48 Then
                ClickMenu GlobalIDControleOperacional, 2, 0
                ClickMenu GlobalIDControleOperacional, 4, 0
                'espera abrir a tela "Imprimir Agendamento" e muda o título
                GlobalIDTelaImprimirAgendamento = ObtemTelaPrincipalporTitulo("Imprimir Agendamento")
                While GlobalIDTelaImprimirAgendamento = 0
                    Sleep 100
                    DoEvents
                    GlobalIDTelaImprimirAgendamento = ObtemTelaPrincipalporTitulo("Imprimir Agendamento")
                Wend
                SendMessageString GlobalIDTelaImprimirAgendamento, WM_SETTEXT, 0, "Escolha o dia e clique em Visualizar"
                tmVeriricaSeControleEstaAberto.Enabled = False
                'tmVerificaseMenuRequerimentosfoiAcionado.Enabled = False
                tmRelaçãodeRequerimentos.Enabled = True
            End If
        End If
    End If

End Sub






Private Sub tmVeriricaSeControleEstaAberto_Timer()
    Dim titletmp As String
    Dim nret As Long
    Dim localMenu As Long
    Dim RelaçãodeMenu As String
    Dim newMenu As Long
    Dim hToolbar20WndClass As Long
    Dim hmsvb_lib_Tollbar As Long
    Dim hPrimeira As Long
    Dim hSegunda As Long
    Dim ClassedoControle As String
    Dim classlength As Long
    Dim IDtelasInternasdoSABI As Long
    Dim IDApagaTela As Long
    Dim contavezes As Long
    'Se ja houver rodando entao sai
    If App.PrevInstance Then
        MsgBox "O aplicativo 'Requerimentos do SABI' já está aberto.", vbCritical, "Requerimentos do SABI"
        End
    End If

    'sai se controle não está ativo
    GlobalIDControleOperacional = ObtemTelaPrincipalporTitulo("SABI - Módulo de Controle Operacional")
    If GlobalIDControleOperacional = 0 Then
        MsgBox "Antes de abrir este aplicativo abra o módulo 'Controle Operacional' do SABI.", vbCritical, "Requerimentos do SABI"
        End
    End If
    'limpa todas telas internas do SABI
        'resta agora as telas externas
    verificaeapaga "Imprimir Agendamento"
    verificaeapaga "Imprimir Escala"
    verificaeapaga "Marcação da Avaliação Social"
    verificaeapaga "Segunda Via de Carta de Exigência"
    verificaeapaga "Pesquisa de Requerente"

    GlobalhMDIClient = 0
    GlobalhMDIClient = FindWindowEx(GlobalIDControleOperacional, 0, "MDIClient", "")
    If GlobalhMDIClient = 0 Then
        MsgBox "Não foi encontrada o indentificador da tela de fundo do Controle Operacional.", vbCritical, "Requerimentos do SABI"
        End
    End If
    
    IDtelasInternasdoSABI = 0
    IDtelasInternasdoSABI = FindWindowEx(GlobalhMDIClient, 0, vbNullString, vbNullString)
    While IDtelasInternasdoSABI <> 0
        If IDtelasInternasdoSABI <> 0 Then IDApagaTela = IDtelasInternasdoSABI
        DoEvents
        Sleep 300
        IDtelasInternasdoSABI = FindWindowEx(GlobalhMDIClient, IDtelasInternasdoSABI, vbNullString, vbNullString)
        If IDApagaTela <> 0 Then SendMessage IDApagaTela, WM_CLOSE, 0, 0
    Wend
    'todas as telas internas foram limpas


    
    
    'acerta a tela
    ColocaTelaControleOperacionanoModoNormal
    ColocaTelaControleOperacionanoModoMaximizado

    
    'Abre tela Pesquisa Avançada
    ClickMenu GlobalIDControleOperacional, 2, 0
    Sleep 3000
    IDtelasInternasdoSABI = 0
    contavezes = 0
    While IDtelasInternasdoSABI = 0
        IDtelasInternasdoSABI = FindWindowEx(GlobalhMDIClient, 0, "ThunderRT6FormDC", "Consulta Requerimento/Benefício")
        Sleep 300
        DoEvents
        contavezes = contavezes + 1
        If contavezes > 20 Then
            MsgBox "Não foi possível abrir tela 'Consulta Requerimento/Benefício'.", vbCritical, "Consulta Requerimento/Benefício"
            End
        End If
    Wend
    GlobalIDTelaConsultaRequerimentoBenefício = 0
    While GlobalIDTelaConsultaRequerimentoBenefício = 0
        Sleep 300
        DoEvents
        GlobalIDTelaConsultaRequerimentoBenefício = achaTelaInternaAtiva("Consulta Requerimento/Benefício")
    Wend
    'abre Imprimir Agendamento
    Sleep 2000
    ClickMenu GlobalIDControleOperacional, 4, 0
    DoEvents
    
    Sleep 1000
    GlobalIDTelaImprimirAgendamento = 0
    contavezes = 0
    While GlobalIDTelaImprimirAgendamento = 0
        GlobalIDTelaImprimirAgendamento = ObtemTelaPrincipalporTitulo("Imprimir Agendamento")
        Sleep 300
        DoEvents
        contavezes = contavezes + 1
        If contavezes > 20 Then
            MsgBox "Não foi possível abrir tela 'Imprimir Agendamento'.", vbCritical, "Imprimir Agendamento"
            End
        End If
    Wend

     'muda o titulo da tela
    SendMessageString GlobalIDTelaImprimirAgendamento, WM_SETTEXT, 0, "Escolha o dia e clique em Visualizar"
    tmVeriricaSeControleEstaAberto = False
    tmRelaçãodeRequerimentos.Enabled = True

End Sub

Private Sub txtRelacionaMenus_Change()

End Sub

Private Sub decodeRequerimentos(texto As String)
    Dim memo As String
    Dim posmedico As Long
    Dim posRequerimento As Long
    Dim posproximapericia As Long
    Dim posdata As Long
    Dim posfimdata As Long
    Dim datamemo As String

    memo = texto
    If memo <> "" Then
        posdata = InStr(1, UCase(memo), "FEIRA")
        posfimdata = InStr(posdata + 1, UCase(memo), Chr(13))
        If posdata > 0 Then
            ConverteData (Mid(memo, posdata, posfimdata - posdata))
            If GlobalDatadosRequerimentos < Format(Date, "yyyymmdd") Then
                MsgBox "O dia escolhido deve ser hoje ou data posterior"
                End
            
            End If
            datamemo = GlobalDatadosRequerimentos
            posdata = InStr(posdata + 10, UCase(memo), "FEIRA")
            While posdata > 0
                posfimdata = InStr(posdata + 1, UCase(memo), Chr(13))
                ConverteData (Mid(memo, posdata, posfimdata - posdata))
                
                If GlobalDatadosRequerimentos <> datamemo Then
                    MsgBox "Somente um dia deve ser escolhido"
                    End
                End If
                posdata = InStr(posdata + 10, UCase(memo), "FEIRA")
            Wend
        Else
            GlobalDatadosRequerimentos = Format(Date, "yyyymmdd")
        End If
        lblRequerimentodoSABI.Caption = "Requerimentos do dia '" & Mid(GlobalDatadosRequerimentos, 7, 2) & "/" & Mid(GlobalDatadosRequerimentos, 5, 2) & "/" & Mid(GlobalDatadosRequerimentos, 1, 4) & "'"
        
    
    
    
        If InStr(1, memo, "Medico") > 0 Then
            'retira info aps
            memo = Mid(memo, InStr(1, memo, "Medico") - 2)
        End If
        While InStr(1, memo, "Medico") > 0
            posmedico = InStr(1, memo, "Medico")
            posRequerimento = InStr(1, memo, "Requerimento")
            posproximapericia = InStr(posRequerimento + 15, memo, Chr(13)) + 2
            If posproximapericia > posRequerimento Then
            
                memo = Mid(memo, 1, posmedico - 2) & Mid(memo, posproximapericia)
            Else
                 memo = ""
            End If
        Wend
    End If
    MontaListadeRequerimentos (memo)
End Sub

Private Sub txtRequerimentos_Change()

End Sub


Private Sub txtResumo_Change()

End Sub

Private Sub txtIDTelaAtiva_Change()

End Sub
