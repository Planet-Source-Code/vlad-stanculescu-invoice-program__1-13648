VERSION 5.00
Begin VB.Form about 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   3480
      Width           =   2055
      Begin VB.CommandButton Command1 
         Caption         =   "Done"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "This program was written by Vlad Stanculescu. The purpose of this program is for me to learn how to control the print function."
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4095
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   240
         Picture         =   "about.frx":000C
         ScaleHeight     =   2415
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Application"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Unload about
 about.Hide
End Sub
