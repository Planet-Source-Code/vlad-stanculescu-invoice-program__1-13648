VERSION 5.00
Begin VB.Form box 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "box.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim TMP
  
  If Left(box.Caption, 3) = "ABN" Then
  information.Text3.Text = Text1.Text
  End If
  
  If Left(box.Caption, 8) = "Business" Then
  information.Text2.Text = Text1.Text
  End If
 box.Hide
 Unload box

End Sub
Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim TMP
 If KeyCode = 13 Then
  
  If Left(box.Caption, 3) = "ABN" Then
  information.Text3.Text = Text1.Text
  End If
  
  If Left(box.Caption, 8) = "Business" Then
  information.Text2.Text = Text1.Text
  End If
  box.Hide
  Unload box
 End If
End Sub
