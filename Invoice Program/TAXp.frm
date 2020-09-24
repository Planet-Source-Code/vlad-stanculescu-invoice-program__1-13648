VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm TAXp 
   BackColor       =   &H80000001&
   Caption         =   "Invoice Application"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9870
   Icon            =   "TAXp.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7230
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Invoice Application - By Vlad Stanculescu"
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Text            =   "1"
            TextSave        =   "1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "F&ile"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "TAXp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
 Load toolbox
 Load information
 toolbox.Show
 information.Show
End Sub

Private Sub MDIForm_Resize()
 If TAXp.WindowState = vbMaximized Then
  TAXp.WindowState = vbNormal
 End If
 If TAXp.WindowState <> vbMinimized Then
    TAXp.Height = 8180
    TAXp.Width = 9990
 End If
End Sub

Private Sub mnuAbout_Click()
 Load about
 about.Show vbModal, TAXp
End Sub

Private Sub mnuExit_Click()
 End
End Sub

Private Sub mnuHelpT_Click()
 Dim TMP
 TMP = App.Path & "\misc\help.html"
 Shell "start " & TMP
End Sub
