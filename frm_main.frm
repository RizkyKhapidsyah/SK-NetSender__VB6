VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frm_main 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      ToolTipText     =   "Exit !"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmd_about 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "About..."
      Top             =   840
      Width           =   255
   End
   Begin MSComctlLib.ImageList img_lst 
      Left            =   960
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":0442
            Key             =   "appIcon"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt_state 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   840
      Width           =   3855
   End
   Begin VB.CommandButton cmd_snd 
      Caption         =   "GO"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txt_msg 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "Text4"
      ToolTipText     =   "Enter your message here"
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txt_exp 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Text3"
      ToolTipText     =   "from"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txt_dst 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "Text2"
      ToolTipText     =   "to"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txt_nb 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Text            =   "Text1"
      ToolTipText     =   "[ 1 ; 5 ]"
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer tmr_anim 
      Left            =   360
      Top             =   1320
   End
   Begin VB.Shape shp_nb 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   3840
      Top             =   120
      Width           =   525
   End
   Begin VB.Shape shp_dst 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   2160
      Top             =   120
      Width           =   540
   End
   Begin VB.Shape shp_exp 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   360
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lbl_nb 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "NB"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl_dst 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl_exp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare _
    Function NetMessageBufferSend _
    Lib "NETAPI32.DLL" _
    (yServer As Any, _
     yToName As Byte, _
     yFromName As Any, _
     yMsg As Byte, _
     ByVal lSize As Long) As Long
     
Private Declare _
   Function GetComputerName _
   Lib "kernel32" _
   Alias "GetComputerNameA" _
     (ByVal lpBuffer As String, _
      nSize As Long) As Long
     
Private Function _
   BroadcastMessage _
     (pDst As String, _
      pExp As String, _
      pMsg As String) As Long
   
   '----- Variables
   Dim lDstName() As Byte
   Dim lExpName() As Byte
   Dim lMsg() As Byte
   Dim lRet As Long
   
   '----- Affectation des valeurs
   lDstName = pDst & vbNullChar
   lExpName = pExp & vbNullChar
   lMsg = pMsg & vbNullChar
   
   '----- Envoi du message
   lRet = NetMessageBufferSend(lExpName(0), lDstName(0), lExpName(0), lMsg(0), UBound(lMsg))
   BroadcastMessage = lRet
End Function

Private Sub cmd_snd_click()
   '----- Variables
   Dim lRet As Long
   Dim lInd As Integer
   
   '----- Test de validité du nombre de messages envoyés
   If (txt_nb.Text < 1) Then
      txt_nb.Text = 1
   Else
      If (txt_nb.Text > 5) Then
         txt_nb.Text = 5
      End If
   End If
   
   '----- Boucle d'envoi du message
   For lInd = 1 To Int(txt_nb.Text)
      txt_state.Text = "  SENDING..."
      DoEvents
      lRet = BroadcastMessage(txt_dst.Text, txt_exp.Text, txt_msg.Text)
      DoEvents
   Next lInd
   
   '----- Traitement postérieur de l'envoi
   Select Case lRet
      Case 0
         txt_state.Text = "  MESSAGE SUCCESSFULLY SENT !"
      Case 53
         txt_state.Text = "  ERROR : BAD NAME !   ( FROM )"
      Case 123
         txt_state.Text = "  ERROR : EMPTY NAME !"
      Case 2273
         txt_state.Text = "  ERROR : BAD NAME !   ( TO )"
   End Select
End Sub

Private Sub Form_Load()
   '----- Variables locales
   Dim zl_tmp As String
   
   '----- Vérification de la version de Windows
   Call z_windows_version_check
   
   '----- Statut
   txt_state.Text = "  STARTING..."
   
   '----- Personalisation de la barre de titre
   Me.Caption = "Z_NetSender"
   Me.Icon = img_lst.ListImages(1).Picture
   
   '----- Taille de la fenêtre
   Me.Height = zc_titleBarHeight
   Me.Width = 5130
   
   '----- Positionnement de l'interface
   Call z_place_controls
   Call z_resize_window
   
   '----- Affectation des valeurs par défaut
   zl_tmp = z_getComputerName
   txt_exp.Text = zl_tmp
   txt_dst.Text = zl_tmp
   txt_msg.Text = "Your message..."
   txt_nb.Text = 1
   
   '----- Paramétrage de l'animation de barre de titre
   gstep = 0
   tmr_anim.Interval = 200
   tmr_anim.Enabled = True
   
   '----- Statut
   txt_state.Text = "  READY"
End Sub

Private Sub z_windows_version_check()
   '----- Variables locales
   Dim zl_rep As Integer
   Dim zl_num As Integer
   '-----
   zl_num = getVersion()
   If (zl_num <> 2) Then
      zl_rep = MsgBox("This application only runs under WindowsNT systems !", vbExclamation, "Z_NetSender : Warning !")
      DoEvents
      zl_rep = MsgBox("This application is going to be closed.", vbExclamation, "Z_NetSender : Warning !")
      DoEvents
      Unload frm_main
   End If
End Sub

Private Sub z_place_controls()
   '----- Variables locales
   Dim zl_marge As Integer
   
   '-----
   frm_main.ScaleMode = vbPixels
   
   '----- Shapes (Cadres)
   shp_exp.Left = zc_marge
   shp_dst.Left = 6 * zc_marge + zc_larg_cadres + zc_larg_txt
   shp_nb.Left = 11 * zc_marge + 2 * zc_larg_cadres + 2 * zc_larg_txt
      shp_exp.Top = zc_marge
      shp_dst.Top = zc_marge
      shp_nb.Top = zc_marge
   shp_exp.Width = zc_larg_cadres
   shp_dst.Width = zc_larg_cadres
   shp_nb.Width = zc_larg_cadres
      shp_exp.Height = zc_haut_cadres
      shp_dst.Height = zc_haut_cadres
      shp_nb.Height = zc_haut_cadres
   
   '----- TextBox (zones de saisie)
   txt_exp.Left = 2 * zc_marge + zc_larg_cadres
   txt_dst.Left = 7 * zc_marge + 2 * zc_larg_cadres + zc_larg_txt
   txt_nb.Left = 12 * zc_marge + 3 * zc_larg_cadres + 2 * zc_larg_txt
   txt_msg.Left = 2 * zc_marge + zc_larg_cadres
      txt_exp.Top = zc_marge
      txt_dst.Top = zc_marge
      txt_nb.Top = zc_marge
      txt_msg.Top = zc_marge + zc_haut_cadres + zc_marge
   txt_exp.Width = zc_larg_txt
   txt_dst.Width = zc_larg_txt
   txt_nb.Width = Int(zc_larg_txt / 3)
   txt_msg.Width = 2 * zc_larg_txt + Int(zc_larg_txt / 3) + 2 * zc_larg_cadres + 10 * zc_marge
   
   '----- Labels
   lbl_exp.Left = zc_marge
   lbl_dst.Left = 6 * zc_marge + zc_larg_cadres + zc_larg_txt
   lbl_nb.Left = 11 * zc_marge + 2 * zc_larg_cadres + 2 * zc_larg_txt
      lbl_exp.Top = zc_marge + 4
      lbl_dst.Top = zc_marge + 4
      lbl_nb.Top = zc_marge + 4
   lbl_exp.Width = zc_larg_cadres
   lbl_dst.Width = zc_larg_cadres
   lbl_nb.Width = zc_larg_cadres
   
   '----- CommandButton (bouton d'envoi)
   cmd_snd.Left = zc_marge
   cmd_snd.Top = 2 * zc_marge + zc_haut_cadres
   cmd_snd.Width = zc_larg_cadres
   cmd_snd.Height = zc_haut_cadres
   
   '----- State TextBox (Barre de statut personnalisée)
   txt_state.Left = 0
   txt_state.Top = 3 * zc_marge + 2 * zc_haut_cadres
   txt_state.Width = 4 * zc_marge + zc_larg_cadres + txt_msg.Width - cmd_exit.Width - cmd_about.Width - 1
   txt_state.Height = zc_haut_state
   
   '----- CommandButton cmd_exit
   cmd_exit.Top = txt_state.Top
   cmd_exit.Left = txt_msg.Left + txt_msg.Width + zc_marge - cmd_exit.Width + 2
   
   '----- CommandButton cmd_about
   cmd_about.Top = txt_state.Top
   cmd_about.Left = cmd_exit.Left - cmd_about.Width + 1
End Sub

Private Sub z_resize_window()
   frm_main.Width = (txt_state.Width + cmd_about.Width + cmd_exit.Width - zc_marge + zc_formWidthOffset) * 15
   frm_main.Height = (3 * zc_marge + 2 * zc_haut_cadres + txt_state.Height + zc_titleBarHeight) * 15
End Sub

Private Sub tmr_anim_Timer()
   Select Case gstep
      Case 0
         Me.Caption = "Z_NetSender |"
      Case 1
         Me.Caption = "Z_NetSender /"
      Case 2
         Me.Caption = "Z_NetSender --"
      Case 3
         Me.Caption = "Z_NetSender \"
   End Select
   gstep = (gstep + 1) Mod 4
End Sub

Private Function z_getComputerName() As String
   '----- Déclaration des variables
   Dim zl_buffer As String * 25
   Dim zl_ret As Long '
   
   '----- Récupération de l'identifiant de l'ordinateur
   zl_ret = GetComputerName(zl_buffer, 25)
   z_getComputerName = Left(zl_buffer, InStr(zl_buffer, Chr(0)) - 1)
End Function

Private Sub txt_msg_gotFocus()
    txt_msg.SelStart = 0
    txt_msg.SelLength = Len(txt_msg.Text)
End Sub

Private Sub cmd_exit_Click()
   Unload frm_main
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   End
End Sub

Private Sub cmd_about_Click()
   txt_state.Text = "Z_NetSender (Build 007) - April 2001 - Programmed by ZOGALT"
End Sub
