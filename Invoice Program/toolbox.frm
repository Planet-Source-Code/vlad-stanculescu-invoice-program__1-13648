VERSION 5.00
Begin VB.Form toolbox 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   5310
   ClientLeft      =   9675
   ClientTop       =   45
   ClientWidth     =   1935
   Icon            =   "toolbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Command4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Picture         =   "toolbox.frx":000C
      ScaleHeight     =   615
      ScaleWidth      =   1815
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox Command3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Picture         =   "toolbox.frx":0A1F
      ScaleHeight     =   615
      ScaleWidth      =   1815
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox Command2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Picture         =   "toolbox.frx":1941
      ScaleHeight     =   615
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox Command1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Picture         =   "toolbox.frx":3244
      ScaleHeight     =   615
      ScaleWidth      =   1815
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   1800
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   24
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "toolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Load box
 box.Label1 = "Please write down your Australian Business Number (ABN) and press 'Done' when you are finished."
 box.Caption = "ABN (Australian Business Number)"
 box.Show vbModal, TAXp
End Sub

Private Sub Command2_Click()
 Load box
 box.Label1 = "Please enter your business address."
 box.Caption = "Business Address"
 box.Show vbModal, TAXp
End Sub

Private Sub Command3_Click()
 Dim T1
 T1 = MsgBox("Are you sure you want to clear?", vbYesNo, "Clear")
 If T1 = vbYes Then
  taxform.Text1.Text = ""
  taxform.Text2.Text = ""
  taxform.Text3.Text = ""
  taxform.Text4.Text = ""
  taxform.Text5.Text = ""
  taxform.Text6.Text = ""
  taxform.Text7.Text = ""
  taxform.Text8.Text = ""
  taxform.Text9.Text = ""
  taxform.Text10.Text = ""
  taxform.Text11.Text = ""
  taxform.Text12.Text = ""
  taxform.Text13.Text = ""
  taxform.Text14.Text = ""
  taxform.Text15.Text = ""
  taxform.Text16.Text = ""
  taxform.Text17.Text = ""
  taxform.Text18.Text = ""
  taxform.Text19.Text = ""
  taxform.Text20.Text = ""
  taxform.Text21.Text = ""
  taxform.Text22.Text = ""
  taxform.Text23.Text = ""
  taxform.Text24.Text = ""
  taxform.Text25.Text = ""
  taxform.Text26.Text = ""
  taxform.Text27.Text = ""
  taxform.Text28.Text = ""
  taxform.Text29.Text = ""
  taxform.Text30.Text = ""
  taxform.Text31.Text = ""
  taxform.Text32.Text = ""
  taxform.Text33.Text = ""
  taxform.Text34.Text = ""
  taxform.Text35.Text = ""
  taxform.Text36.Text = ""
  taxform.Text37.Text = ""
  taxform.Text38.Text = ""
  taxform.Text39.Text = ""
  taxform.Text40.Text = ""
  taxform.Text41.Text = ""
  information.Text1.Text = ""
 End If
End Sub

Private Sub Command4_Click()
 If information.Text1.Text <> "" And information.Text3.Text <> "" And taxform.Text41.Text <> "" Then
 
 Dim Answer As String, HorizontalMargin, VerticalMargin As Single
 Dim MyCenteredText As String, MyCenteredTextWidth As Single
 Dim MyLeftText As String, MyLeftTextWidth As Single
 Dim MyRightText As String, MyRightTextWidth As Single
 Dim txtGrid15 As String
 Dim txtGrid11, txtGrid12, txtGrid13, txtGrid14, txtGrid21, txtGrid22 As String
 Dim txtGrid23, txtGrid24, txtGrid31, txtGrid32, txtGrid33, txtGrid34 As String
 Dim MyGridTitle As String, MyGridTitleWidth As Single
 Dim Row1Col1Left, Row1Col2Left, Row1Col3Left, Row1Col4Left As Single
 Dim Row2Col1Left, Row2Col2Left, Row2Col3Left, Row2Col4Left As Single
 Dim Row3Col1Left, Row3Col2Left, Row3Col3Left, Row3Col4Left As Single
 Dim Row1Top, Row2Top, Row3Top, ImageLeft, ImageTop As Single

 Printer.ScaleMode = vbCentimeters
 'HorizontalMargin = (21 - Printer.ScaleWidth) / 2
 'VerticalMargin = (29.7 - Printer.ScaleHeight) / 2
 'HorizontalMargin = 1 + HorizontalMargin
 'VerticalMargin = 1.5 + VerticalMargin
 Printer.Print "";
 'Printer.Line (HorizontalMargin, VerticalMargin)-(21 - HorizontalMargin, 29.7 - VerticalMargin), RGB(255, 0, 0), B
 
 Printer.FontName = "FixedSys"
 Printer.FontSize = 24
 Printer.FontBold = True          'we want bold
 Printer.FontItalic = False       'no italic
 Printer.FontUnderline = False    'no underline
 Printer.FontStrikethru = False   'no strike
 Printer.ForeColor = RGB(0, 0, 0) 'color black
 
 MyCenteredText = "TAX INVOICE/STATEMENT"
 MyCenteredTextWidth = Printer.TextWidth(MyCenteredText)
 Printer.CurrentX = (21 - MyCenteredTextWidth) / 2
 Printer.CurrentY = VerticalMargin + 0.5
 Printer.Print MyCenteredText
 
 Printer.FontName = "FixedSys"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.FontItalic = False
 Printer.FontUnderline = False
 Printer.FontStrikethru = Fals
 Printer.ForeColor = RGB(0, 0, 0)

 MyLeftText = "Time           : " & Format(Date, "Long Date")
 MyLeftTextWidth = Printer.TextWidth(MyLeftText)
 If MyLeftTextWidth > 21 - (HorizontalMargin * 2) Then Exit Sub
 Printer.CurrentX = 1
 Printer.CurrentY = VerticalMargin + "2.0"
 Printer.Print MyLeftText
 
 MyLeftText = "To             : " & information.Text1.Text
 MyLeftTextWidth = Printer.TextWidth(MyLeftText)
 If MyLeftTextWidth > 21 - (HorizontalMargin * 2) Then Exit Sub
 Printer.CurrentX = 1
 Printer.CurrentY = VerticalMargin + "2.5"
 Printer.Print MyLeftText
 
 MyLeftText = "From           : " & information.Text2.Text
 MyLeftTextWidth = Printer.TextWidth(MyLeftText)
 If MyLeftTextWidth > 21 - (HorizontalMargin * 2) Then Exit Sub
 Printer.CurrentX = 1
 Printer.CurrentY = VerticalMargin + "3.0"
 Printer.Print MyLeftText
 
 MyLeftText = "ABN            : " & information.Text3.Text
 MyLeftTextWidth = Printer.TextWidth(MyLeftText)
 If MyLeftTextWidth > 21 - (HorizontalMargin * 2) Then Exit Sub
 Printer.CurrentX = 1
 Printer.CurrentY = VerticalMargin + "3.5"
 Printer.Print MyLeftText
 MyLeftText = "To             : " & information.Text1.Text
 MyLeftTextWidth = Printer.TextWidth(MyLeftText)
 If MyLeftTextWidth > 21 - (HorizontalMargin * 2) Then Exit Sub
 Printer.CurrentX = 1
 Printer.CurrentY = VerticalMargin + "4.0"
 Printer.Print MyLeftText
 
 Printer.CurrentX = 1
 Printer.CurrentY = VerticalMargin + "5"
 Printer.FontName = "FixedSys"
 Printer.FontSize = 9
 Printer.FontBold = False
 Printer.FontItalic = False
 Printer.FontUnderline = False
 Printer.FontStrikethru = Fals
 Printer.ForeColor = RGB(0, 0, 0)
 Printer.CurrentX = 1
 Printer.Print "Total fee      : " + taxform.Text41.Text
 Printer.Print Space(1)
 Printer.CurrentX = 1
 Printer.Print Space(1)
 Printer.CurrentX = 1
 Printer.FontUnderline = True
 Printer.Print "ID  " + "Item(s)  " + "Description                        " + "Price     " + "Total Balance"
 Printer.FontUnderline = False
 Printer.CurrentX = 1
 Printer.Print LST(1)
 Printer.CurrentX = 1
 Printer.Print LST(2)
 Printer.CurrentX = 1
 Printer.Print LST(3)
 Printer.CurrentX = 1
 Printer.Print LST(4)
 Printer.CurrentX = 1
 Printer.Print LST(5)
 Printer.CurrentX = 1
 Printer.Print LST(6)
 Printer.CurrentX = 1
 Printer.Print LST(7)
 Printer.CurrentX = 1
 Printer.Print LST(8)
 Printer.CurrentX = 1
 Printer.Print LST(9)
 Printer.CurrentX = 1
 Printer.Print LST(10)
 Printer.CurrentX = 1
 Printer.EndDoc
 Exit Sub
 End If
 TMP = MsgBox("You must fill in ALL details!", vbInformation + vbOKOnly, "Missing Information")
End Sub

Private Sub Command5_Click()
 Load box
 box.Label1 = "Please enter the Goods & Services Tax as a percentage, please leave out the %. (Default: 10%)"
 box.Caption = "GST (Goods & Services Tax)"
 box.Show vbModal, TAXp
End Sub

Private Sub Command6_Click()
 Load box
 box.Label1 = "Please enter in the business/tax fee that you would like to add to each price in dollars. (Example: $2.00)"
 box.Caption = "Added Fee"
 box.Show vbModal, TAXp
End Sub

Private Sub Form_Load()
 toolbox.Top = 50
 toolbox.Left = 50
End Sub
