Attribute VB_Name = "mod2"
Option Explicit
Function LST(NUM As Integer)
 Dim TMP As String, INF1, INF2, INF3, INF4, INF5
 If NUM = 1 Then
 If taxform.Text4.Text <> "" Then
  INF1 = "1.  "
  INF2 = taxform.Text1.Text + Space((9 - (Len(taxform.Text1.Text))))
  INF3 = taxform.Text2.Text + Space((35 - (Len(taxform.Text2.Text))))
  INF4 = Format(taxform.Text3.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text4.Text + Space((10 - (Len(taxform.Text4.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If

 If NUM = 2 Then
 If taxform.Text5.Text <> "" Then
  INF1 = "2.  "
  INF2 = taxform.Text8.Text + Space((9 - (Len(taxform.Text8.Text))))
  INF3 = taxform.Text7.Text + Space((35 - (Len(taxform.Text7.Text))))
  INF4 = Format(taxform.Text6.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text5.Text + Space((10 - (Len(taxform.Text5.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If

 If NUM = 3 Then
 If taxform.Text12.Text <> "" Then
  INF1 = "3.  "
  INF2 = taxform.Text9.Text + Space((9 - (Len(taxform.Text9.Text))))
  INF3 = taxform.Text10.Text + Space((35 - (Len(taxform.Text10.Text))))
  INF4 = Format(taxform.Text11.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text12.Text + Space((10 - (Len(taxform.Text12.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If

 If NUM = 4 Then
 If taxform.Text13.Text <> "" Then
  INF1 = "4.  "
  INF2 = taxform.Text16.Text + Space((9 - (Len(taxform.Text16.Text))))
  INF3 = taxform.Text15.Text + Space((35 - (Len(taxform.Text15.Text))))
  INF4 = Format(taxform.Text14.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text13.Text + Space((10 - (Len(taxform.Text13.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If

 If NUM = 5 Then
 If taxform.Text20.Text <> "" Then
  INF1 = "5.  "
  INF2 = taxform.Text17.Text + Space((9 - (Len(taxform.Text17.Text))))
  INF3 = taxform.Text18.Text + Space((35 - (Len(taxform.Text18.Text))))
  INF4 = Format(taxform.Text19.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text20.Text + Space((10 - (Len(taxform.Text20.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If

 If NUM = 6 Then
 If taxform.Text21.Text <> "" Then
  INF1 = "6.  "
  INF2 = taxform.Text24.Text + Space((9 - (Len(taxform.Text24.Text))))
  INF3 = taxform.Text23.Text + Space((35 - (Len(taxform.Text23.Text))))
  INF4 = Format(taxform.Text22.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text21.Text + Space((10 - (Len(taxform.Text21.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If

 If NUM = 7 Then
 If taxform.Text25.Text <> "" Then
  INF1 = "7.  "
  INF2 = taxform.Text28.Text + Space((9 - (Len(taxform.Text28.Text))))
  INF3 = taxform.Text27.Text + Space((35 - (Len(taxform.Text27.Text))))
  INF4 = Format(taxform.Text26.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text25.Text + Space((10 - (Len(taxform.Text25.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If

 If NUM = 8 Then
 If taxform.Text32.Text <> "" Then
  INF1 = "8.  "
  INF2 = taxform.Text29.Text + Space((9 - (Len(taxform.Text29.Text))))
  INF3 = taxform.Text30.Text + Space((35 - (Len(taxform.Text30.Text))))
  INF4 = Format(taxform.Text31.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text32.Text + Space((10 - (Len(taxform.Text32.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If

 If NUM = 9 Then
 If taxform.Text36.Text <> "" Then
  INF1 = "9.  "
  INF2 = taxform.Text33.Text + Space((9 - (Len(taxform.Text33.Text))))
  INF3 = taxform.Text34.Text + Space((35 - (Len(taxform.Text34.Text))))
  INF4 = Format(taxform.Text35.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text36.Text + Space((10 - (Len(taxform.Text36.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If

 If NUM = 10 Then
 If taxform.Text40.Text <> "" Then
  INF1 = "10. "
  INF2 = taxform.Text37.Text + Space((9 - (Len(taxform.Text37.Text))))
  INF3 = taxform.Text38.Text + Space((35 - (Len(taxform.Text38.Text))))
  INF4 = Format(taxform.Text39.Text, "$###,###.##")
  If Left(Right(INF4, 2), 1) = "." Then INF4 = INF4 + "0"
  If Right(INF4, 1) = "." Then INF4 = INF4 + "00"
  INF4 = INF4 + Space((10 - (Len(INF4))))
  INF5 = taxform.Text40.Text + Space((10 - (Len(taxform.Text40.Text))))
  LST = INF1 + INF2 + INF3 + INF4 + INF5
 End If
 End If
End Function

