VERSION 5.00
Begin VB.Form taxform 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "3"
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   Icon            =   "taxform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   120
      TabIndex        =   46
      Top             =   1680
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   1320
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6480
      TabIndex        =   39
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5280
      TabIndex        =   38
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   720
      TabIndex        =   37
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   120
      TabIndex        =   36
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   720
      TabIndex        =   35
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   5280
      TabIndex        =   34
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   6480
      TabIndex        =   33
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   6480
      TabIndex        =   32
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   5280
      TabIndex        =   31
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   720
      TabIndex        =   30
      Top             =   2400
      Width           =   4455
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   120
      TabIndex        =   29
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   720
      TabIndex        =   27
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   5280
      TabIndex        =   26
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   6480
      TabIndex        =   25
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   6480
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   5280
      TabIndex        =   23
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   720
      TabIndex        =   22
      Top             =   3120
      Width           =   4455
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   6480
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text26 
      Height          =   285
      Left            =   5280
      TabIndex        =   19
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text27 
      Height          =   285
      Left            =   720
      TabIndex        =   18
      Top             =   3480
      Width           =   4455
   End
   Begin VB.TextBox Text28 
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text29 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox Text30 
      Height          =   285
      Left            =   720
      TabIndex        =   15
      Top             =   3840
      Width           =   4455
   End
   Begin VB.TextBox Text31 
      Height          =   285
      Left            =   5280
      TabIndex        =   14
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text32 
      Height          =   285
      Left            =   6480
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text33 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox Text34 
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Top             =   4200
      Width           =   4455
   End
   Begin VB.TextBox Text35 
      Height          =   285
      Left            =   5280
      TabIndex        =   10
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text37 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox Text38 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   4560
      Width           =   4455
   End
   Begin VB.TextBox Text39 
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text41 
      Height          =   285
      Left            =   6480
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text36 
      Height          =   285
      Left            =   6480
      TabIndex        =   5
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text40 
      Height          =   285
      Left            =   6480
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6480
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Form Entry Page"
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
      TabIndex        =   45
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   43
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   42
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   41
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   40
      Top             =   4920
      Width           =   1095
   End
End
Attribute VB_Name = "taxform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Timer1.Interval = 50
 Timer1.Enabled = True
 taxform.Top = 50
 taxform.Left = 2050
End Sub

Function CalcFields()
 Dim T, GST As Integer, TMP5, TMP6
 If Text1.Text = "" Or Text3.Text = "" Then Text4.Text = ""
 If Text1.Text <> "" And Text3.Text <> "" Then
    T = (Text1.Text * Text3.Text)
    Text4.Text = "$" & Format(T, "###,###.#######")
    If Right(Text4.Text, 1) = "." Then Text4.Text = Text4.Text + "00"
 End If
 
 If Text8.Text = "" Or Text6.Text = "" Then Text5.Text = ""
 If Text8.Text <> "" And Text6.Text <> "" Then
    T = (Text8.Text * Text6.Text)
    Text5.Text = "$" & Format(T, "###,###.#######")
    If Right(Text5.Text, 1) = "." Then Text5.Text = Text5.Text + "00"
 End If
 
 If Text9.Text = "" Or Text11.Text = "" Then Text12.Text = ""
 If Text9.Text <> "" And Text11.Text <> "" Then
    T = (Text9.Text * Text11.Text)
    Text12.Text = "$" & Format(T, "###,###.#######")
    If Right(Text12.Text, 1) = "." Then Text12.Text = Text12.Text + "00"
 End If
 
 If Text16.Text = "" Or Text14.Text = "" Then Text13.Text = ""
 If Text16.Text <> "" And Text14.Text <> "" Then
    T = (Text16.Text * Text14.Text)
    Text13.Text = "$" & Format(T, "###,###.#######")
    If Right(Text13.Text, 1) = "." Then Text13.Text = Text13.Text + "00"
 End If
 
 If Text17.Text = "" Or Text19.Text = "" Then Text20.Text = ""
 If Text17.Text <> "" And Text19.Text <> "" Then
    T = (Text17.Text * Text19.Text)
    Text20.Text = "$" & Format(T, "###,###.#######")
    If Right(Text20.Text, 1) = "." Then Text20.Text = Text20.Text + "00"
 End If
 
 If Text24.Text = "" Or Text22.Text = "" Then Text21.Text = ""
 If Text24.Text <> "" And Text22.Text <> "" Then
    T = (Text24.Text * Text22.Text)
    Text21.Text = "$" & Format(T, "###,###.#######")
    If Right(Text21.Text, 1) = "." Then Text21.Text = Text21.Text + "00"
 End If
 
 If Text28.Text = "" Or Text26.Text = "" Then Text25.Text = ""
 If Text28.Text <> "" And Text26.Text <> "" Then
    T = (Text28.Text * Text26.Text)
    Text25.Text = "$" & Format(T, "###,###.#######")
    If Right(Text25.Text, 1) = "." Then Text25.Text = Text25.Text + "00"
 End If
 
 If Text29.Text = "" Or Text31.Text = "" Then Text32.Text = ""
 If Text29.Text <> "" And Text31.Text <> "" Then
    T = (Text29.Text * Text31.Text)
    Text32.Text = "$" & Format(T, "###,###.#######")
    If Right(Text32.Text, 1) = "." Then Text32.Text = Text32.Text + "00"
 End If
 
 If Text33.Text = "" Or Text35.Text = "" Then Text36.Text = ""
 If Text33.Text <> "" And Text35.Text <> "" Then
    T = (Text33.Text * Text35.Text)
    Text36.Text = "$" & Format(T, "###,###.#######")
    If Right(Text36.Text, 1) = "." Then Text36.Text = Text36.Text + "00"
 End If
 
 If Text37.Text = "" Or Text39.Text = "" Then Text40.Text = ""
 If Text37.Text <> "" And Text39.Text <> "" Then
    T = (Text37.Text * Text39.Text)
    Text40.Text = "$" & Format(T, "###,###.#######")
    If Right(Text40.Text, 1) = "." Then Text40.Text = Text40.Text + "00"
 End If
End Function
Private Sub Timer1_Timer()
 CalcFields
 CalcEnd
 FixFields
End Sub
Function CalcEnd()
 Dim TMP
 If Text4.Text <> "" Then TMP = (TMP + Val(Format(Text4.Text, "#.##")))
 If Text5.Text <> "" Then TMP = (TMP + Val(Format(Text5.Text, "#.##")))
 If Text12.Text <> "" Then TMP = (TMP + Val(Format(Text12.Text, "#.##")))
 If Text13.Text <> "" Then TMP = (TMP + Val(Format(Text13.Text, "#.##")))
 If Text20.Text <> "" Then TMP = (TMP + Val(Format(Text20.Text, "#.##")))
 If Text21.Text <> "" Then TMP = (TMP + Val(Format(Text21.Text, "#.##")))
 If Text25.Text <> "" Then TMP = (TMP + Val(Format(Text25.Text, "#.##")))
 If Text32.Text <> "" Then TMP = (TMP + Val(Format(Text32.Text, "#.##")))
 If Text36.Text <> "" Then TMP = (TMP + Val(Format(Text36.Text, "#.##")))
 If Text40.Text <> "" Then TMP = (TMP + Val(Format(Text40.Text, "#.##")))
 If TMP <> "0" Or TMP <> "" Then Text41.Text = Format(TMP, "$###,###.##")
End Function
Function FixFields()
 If Len(Format(Text4.Text, "$###,###.#0")) <> Len(Text4.Text) Then Text4.Text = Format(Text4.Text, "$###,###.#0")
 If Len(Format(Text5.Text, "$###,###.#0")) <> Len(Text5.Text) Then Text5.Text = Format(Text5.Text, "$###,###.#0")
 If Len(Format(Text12.Text, "$###,###.#0")) <> Len(Text12.Text) Then Text12.Text = Format(Text12.Text, "$###,###.#0")
 If Len(Format(Text13.Text, "$###,###.#0")) <> Len(Text13.Text) Then Text13.Text = Format(Text13.Text, "$###,###.#0")
 If Len(Format(Text20.Text, "$###,###.#0")) <> Len(Text20.Text) Then Text20.Text = Format(Text20.Text, "$###,###.#0")
 If Len(Format(Text21.Text, "$###,###.#0")) <> Len(Text21.Text) Then Text21.Text = Format(Text21.Text, "$###,###.#0")
 If Len(Format(Text25.Text, "$###,###.#0")) <> Len(Text25.Text) Then Text25.Text = Format(Text25.Text, "$###,###.#0")
 If Len(Format(Text32.Text, "$###,###.#0")) <> Len(Text32.Text) Then Text32.Text = Format(Text32.Text, "$###,###.#0")
 If Len(Format(Text36.Text, "$###,###.#0")) <> Len(Text36.Text) Then Text36.Text = Format(Text36.Text, "$###,###.#0")
 If Len(Format(Text40.Text, "$###,###.#0")) <> Len(Text40.Text) Then Text40.Text = Format(Text40.Text, "$###,###.#0")
 If Len(Format(Text41.Text, "$###,###.#0")) <> Len(Text41.Text) Then Text41.Text = Format(Text41.Text, "$###,###.#0")
End Function
Private Sub Text1_Change()
 Static strSaved As String
 If Text1.Text <> "" And Text1.Text <> "-" And Text1.Text <> "." And Text1.Text <> "-." Then
    If Not IsNumeric(Text1.Text) Then Text1.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text1.Text
End Sub
Private Sub Text8_Change()
 Static strSaved As String
 If Text8.Text <> "" And Text8.Text <> "-" And Text8.Text <> "." And Text8.Text <> "-." Then
    If Not IsNumeric(Text8.Text) Then Text8.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text8.Text
End Sub
Private Sub Text9_Change()
 Static strSaved As String
 If Text9.Text <> "" And Text9.Text <> "-" And Text9.Text <> "." And Text9.Text <> "-." Then
    If Not IsNumeric(Text9.Text) Then Text9.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text9.Text
End Sub
Private Sub Text16_Change()
 Static strSaved As String
 If Text16.Text <> "" And Text16.Text <> "-" And Text16.Text <> "." And Text16.Text <> "-." Then
    If Not IsNumeric(Text16.Text) Then Text16.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text16.Text
End Sub
Private Sub Text17_Change()
 Static strSaved As String
 If Text17.Text <> "" And Text17.Text <> "-" And Text17.Text <> "." And Text17.Text <> "-." Then
    If Not IsNumeric(Text17.Text) Then Text17.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text17.Text
End Sub
Private Sub Text24_Change()
 Static strSaved As String
 If Text24.Text <> "" And Text24.Text <> "-" And Text24.Text <> "." And Text24.Text <> "-." Then
    If Not IsNumeric(Text24.Text) Then Text24.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text24.Text
End Sub
Private Sub Text28_Change()
 Static strSaved As String
 If Text28.Text <> "" And Text28.Text <> "-" And Text28.Text <> "." And Text28.Text <> "-." Then
    If Not IsNumeric(Text28.Text) Then Text28.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text28.Text
End Sub
Private Sub Text29_Change()
 Static strSaved As String
 If Text29.Text <> "" And Text29.Text <> "-" And Text29.Text <> "." And Text29.Text <> "-." Then
    If Not IsNumeric(Text29.Text) Then Text29.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text29.Text
End Sub
Private Sub Text33_Change()
 Static strSaved As String
 If Text33.Text <> "" And Text33.Text <> "-" And Text33.Text <> "." And Text33.Text <> "-." Then
    If Not IsNumeric(Text33.Text) Then Text33.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text33.Text
End Sub
Private Sub Text37_Change()
 Static strSaved As String
 If Text37.Text <> "" And Text37.Text <> "-" And Text37.Text <> "." And Text37.Text <> "-." Then
    If Not IsNumeric(Text37.Text) Then Text37.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text37.Text
End Sub
Private Sub Text3_Change()
 Static strSaved As String
 If Text3.Text <> "" And Text3.Text <> "-" And Text3.Text <> "." And Text3.Text <> "-." Then
    If Not IsNumeric(Text3.Text) Then Text3.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text3.Text
End Sub
Private Sub Text6_Change()
 Static strSaved As String
 If Text6.Text <> "" And Text6.Text <> "-" And Text6.Text <> "." And Text6.Text <> "-." Then
    If Not IsNumeric(Text6.Text) Then Text6.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text6.Text
End Sub
Private Sub Text11_Change()
 Static strSaved As String
 If Text11.Text <> "" And Text11.Text <> "-" And Text11.Text <> "." And Text11.Text <> "-." Then
    If Not IsNumeric(Text11.Text) Then Text11.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text11.Text
End Sub
Private Sub Text14_Change()
 Static strSaved As String
 If Text14.Text <> "" And Text14.Text <> "-" And Text14.Text <> "." And Text14.Text <> "-." Then
    If Not IsNumeric(Text14.Text) Then Text14.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text14.Text
End Sub
Private Sub Text19_Change()
 Static strSaved As String
 If Text19.Text <> "" And Text19.Text <> "-" And Text19.Text <> "." And Text19.Text <> "-." Then
    If Not IsNumeric(Text19.Text) Then Text19.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text19.Text
End Sub
Private Sub Text22_Change()
 Static strSaved As String
 If Text22.Text <> "" And Text22.Text <> "-" And Text22.Text <> "." And Text22.Text <> "-." Then
    If Not IsNumeric(Text22.Text) Then Text22.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text22.Text
End Sub
Private Sub Text26_Change()
 Static strSaved As String
 If Text26.Text <> "" And Text26.Text <> "-" And Text26.Text <> "." And Text26.Text <> "-." Then
    If Not IsNumeric(Text26.Text) Then Text26.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text26.Text
End Sub
Private Sub Text31_Change()
 Static strSaved As String
 If Text31.Text <> "" And Text31.Text <> "-" And Text31.Text <> "." And Text31.Text <> "-." Then
    If Not IsNumeric(Text31.Text) Then Text31.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text31.Text
End Sub
Private Sub Text35_Change()
 Static strSaved As String
 If Text35.Text <> "" And Text35.Text <> "-" And Text35.Text <> "." And Text35.Text <> "-." Then
    If Not IsNumeric(Text35.Text) Then Text35.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text35.Text
End Sub
Private Sub Text39_Change()
 Static strSaved As String
 If Text39.Text <> "" And Text39.Text <> "-" And Text39.Text <> "." And Text39.Text <> "-." Then
    If Not IsNumeric(Text39.Text) Then Text39.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text39.Text
End Sub
