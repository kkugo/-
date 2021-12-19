Attribute VB_Name = "Module1"


Const Ltofile = 1
Const Lfromdata = 4
Const Lcheck = 2
Const R678 = 4
Const Rgyosha = 5
Const Rmnum = 6
Const Rshina = 7
Const Rtpm = 10
Const Rfutan = 11
Const Rringi = 12
Const Rkanko = 13
Const Rkansht = 14
Const Rsuryo = 15
Const Runit = 16
Const Rtanka = 17
Const Rzei = 18
Const Rdataend = 19
Const D678 = 12
Const Dgyosha = 1
Const Dmnum = 3
Const Dshina = 4
Const Dfutan = 19
Const Dkanko = 15
Const Dkansht = 17
Const Dsuryo = 5
Const Dunit = 6
Const Dtanka = 7

Function autozeiritsu(shi As Integer, hin As String) As Integer
    Dim i As Integer
    Dim sh As Worksheet
    Set sh = Worksheets("Sheet2")
    hin = StrConv(hin, vbNarrow)
      i = 2
      autozeiritsu = 10
      While sh.Cells(i, 5) <> ""
        If sh.Cells(i, 5).Value = shi And InStr(hin, sh.Cells(i, 6).Text) > 0 Then
            autozeiritsu = sh.Cells(i, 7).Value
            Exit Function
        End If
        i = i + 1
      Wend

End Function


Sub line2file()
Attribute line2file.VB_ProcData.VB_Invoke_Func = "a\n14"

   Dim databak As Variant
   Dim datarsl As Variant
   Dim strtmp As String
   
   '自動計算・画面更新OFF
   Application.ScreenUpdating = False
   Application.Calculation = xlCalculationManual
   
   '前にファイルコピーしたデータをバックアップ
   databak = Range(Cells(Ltofile, 1), Cells(Ltofile, Rdataend))
   
   'ファイル書き出し↓
    wc = Asc("A")
    'ファイルオープン
    fldn = "C:\localprg\uwsc5302\uws\"
    fn = fldn + "ex.txt"
    Open fn For Output As #1
     
    '選択行を上にコピー
    j = ActiveCell.Row
    For i = 1 To Rdataend
            Cells(Lfromdata, i) = Cells(j, i)
    Next
    
    'チェックに応じて品名を作成
    strtmp = ""
    For i = 0 To 2
     If Cells(Lcheck, Rshina + i) = True Then
        strtmp = strtmp & Cells(j, Rshina + i)
     End If
    Next
    Cells(Lfromdata, Rshina) = strtmp
    Cells(Lfromdata, Rshina + 1) = ""
    Cells(Lfromdata, Rshina + 2) = ""
    'チェックに応じて品名作成↑
    
    '値引・送料
    If InStr(Cells(Lfromdata, Rshina).Text, "値引") > 0 Then
        Cells(Lfromdata, Rsuryo) = 0
        Cells(Lfromdata, Rtanka) = 0
    End If
    If InStr(Cells(Lfromdata, Rshina).Text, "送料") > 0 Then
        Cells(Lfromdata, R678) = 8
        Cells(Lfromdata, Rkanko) = 6570
        Cells(Lfromdata, Rkansht) = 3
        Cells(Lfromdata, Rsuryo) = 1
    End If
    '税率を10 or 8 自動
    Cells(Lfromdata, Rzei) = autozeiritsu(Cells(Lfromdata, Rgyosha).Value, Cells(Lfromdata, Rshina).Text)
    
    '実際に書き出し↓
    For i = 0 To Rdataend - 1
        If i < Rshina Or i > Rtpm - 1 Then
                If Cells(Ltofile, i + 1).Text <> "" Then
                        Print #1, Cells(Ltofile, i + 1).Text
                Else:
                        Print #1, ""
                End If
        End If
    Next
    Close #1
    '実際に書きだし↑
    'ファイル書き出し↑
    
    '選択行の下を消す
    r = ActiveCell.Row
    Range(Cells(r + 1, 1), Cells(r + 4, Rdataend)) = ""
    
   '自動計算・画面更新ON
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    'データが変わってなかったらファルコン呼び出す
    datarsl = Range(Cells(Ltofile, 1), Cells(Ltofile, Rdataend))
        
    For i = 1 To Rdataend
        If Format(databak(1, i)) <> Format(datarsl(1, i)) Then Exit Sub
    Next
    
    AppActivate ("FALCON - ")
    
End Sub
Sub addnextline()
Attribute addnextline.VB_ProcData.VB_Invoke_Func = "z\n14"
    j = ActiveCell.Row
    For i = 1 To Rdataend
      If Cells(j + 1, i) = "" Then
         Cells(j + 1, i) = Cells(j, i)
      End If
    Next
End Sub
Function ishankaku(st As String) As Boolean
    Dim h, s As String
    s = Left(st, 1)
    h = StrConv(s, vbWide)
    
    If s = h Then ishankaku = False Else ishankaku = True
    

End Function

Function addtpm(st As String, tp As String) As String
    Dim mojisu As Integer
    mojisu = 0
    l = Len(st)
    fl = InStr("tTkKyYｔＴｋKｙY", tp)
    
    If Len(tp) = 0 Or fl = 0 Then
      addtpm = st + ("                               ")
      Exit Function
    End If
    If ishankaku(Mid(st, 1, 1)) Then mojisu = 1 Else mojisu = 3
    i = 2
    While mojisu < 26 And i <= l
        If ishankaku(Mid(st, i, 1)) Then
            mojisu = mojisu + 1
            If ishankaku(Mid(st, i - 1, 1)) = False Then mojisu = mojisu + 1
        Else
            mojisu = mojisu + 2
            If ishankaku(Mid(st, i - 1, 1)) = True Then mojisu = mojisu + 1
        End If
        i = i + 1
    Wend
    st = Left(st, i - 1)
    If ishankaku(Right(st, 1)) = False Then
       mojisu = mojisu + 1
    End If
    If mojisu < 27 Then
       st = st + Left("                                 ", 27 - mojisu)
    End If
    
    addtpm = st + "tpm-" + tp

End Function

Function nzleft(ByVal st As String, l As Integer) As String
  If l = 0 Then
    nzleft = ""
  Else
    nzleft = Left(st, l)
  End If
End Function

Function nzright(ByVal st As String, l As Integer) As String
  If l = 0 Then
    nzright = ""
  Else
    nzright = Right(st, l)
  End If
End Function

Sub copyfromexam()
Attribute copyfromexam.VB_ProcData.VB_Invoke_Func = "b\n14"

    If Selection.Areas.Count <> 2 Then Exit Sub
    
    c0r = Selection.Areas(1).Column
    c0l = Selection.Areas(1).Row
    c1r = Selection.Areas(2).Column
    c1l = Selection.Areas(2).Row
    
    
'下のセルから上へコピーするようにする
    If c1l < c0l Then
        t = c1l
        c1l = c0l
        c0l = t
        t = c1r
        c1r = c0r
        c0r = t
    End If
    
'日付と消費税を一行前からコピー
    prcR = Array(1, 2, 3, 18)
    For Each r In prcR
        If Cells(c0l, r) = "" Then
          Cells(c0l, r) = Cells(c0l - 1, r)
        End If
    Next

        
'その他を下の選択セルからコピー
    prcR = Array(4, 5, 11, 13, 14, 15, 16, 17)
    
    For Each r In prcR
        Cells(c0l, r) = Cells(c1l, r)
    Next
    
'品名文字列をいろいろ処理する
    sttmp = Cells(c1l, Rshina)

'？？月を自動で入れ替える
    x0 = InStr(sttmp, "月")
    x1 = InStr(sttmp, "ｶﾞﾂ")
    x2 = InStr(sttmp, "/")
    
    ' / の後が数字じゃなかったら日付ではない
    If x2 < Len(sttmp) Then
       If IsNumeric(Mid(sttmp, x2 + 1, 1)) = False Then
         x2 = 0
       End If
    Else
         x2 = 0
    End If
    ' 都ミゾグチママダは/を日付にしない　パイプがある
    If Cells(c0l, Rgyosha) = 8971 Or Cells(c0l, Rgyosha) = 9573 Or Cells(c0l, Rgyosha) = 8984 Then
        x2 = 0
    End If
    
    ' 月が2文字目以降なら最優先で選ぶ
    If x0 < 2 Then
      If x1 > x0 Then x0 = x1
      If x2 > x0 Then x0 = x2
    End If
    
    If x0 = 2 Then
      If IsNumeric(Left(sttmp, 1)) = True Then
        sttmp = Format(Cells(c0l, 2)) + nzright(sttmp, Len(sttmp) - 1)
      End If
    ElseIf x0 > 2 Then
      If IsNumeric(Mid(sttmp, x0 - 2, 2)) = True Then
        sttmp = nzleft(sttmp, x0 - 3) + Format(Cells(c0l, 2)) + nzright(sttmp, Len(sttmp) - x0 + 1)
      ElseIf IsNumeric(Mid(sttmp, x0 - 1, 1)) = True Then
        sttmp = nzleft(sttmp, x0 - 2) + Format(Cells(c0l, 2)) + nzright(sttmp, Len(sttmp) - x0 + 1)
    
      End If
    End If
    
'TPM を分離する
    If Len(sttmp) > 4 Then
        If Mid(sttmp, Len(sttmp) - 4, 3) = "TPM" Then
            Cells(c0l, Rshina) = Left(sttmp, Len(sttmp) - 5)
            Cells(c0l, Rtpm) = Right(sttmp, 1)
        Else
          Cells(c0l, Rshina) = sttmp
          Cells(c0l, Rtpm) = ""
        End If
    Else
          Cells(c0l, Rshina) = sttmp
          Cells(c0l, Rtpm) = ""
   End If
    
   Cells(c0l, Rshina + 1) = ""
   Cells(c0l, Rshina + 2) = ""
   
    
    If c0r = Rmnum Then
      Cells(c0l, Rmnum) = Cells(c0l - 1, Rmnum)
    Else
      Cells(c0l, Rmnum) = Cells(c1l, Rmnum)
    End If
 'M～入っていたら負担部署は空白
    If Cells(c0l, Rmnum) <> "" Then Cells(c0l, Rfutan) = ""
    
    Cells(2, 7) = True
    Cells(2, 8) = False
    Cells(2, 9) = False
   
End Sub


Function mojisu(st As String) As Integer
    mojisu = Len(st)
    l = mojisu
    If LenB(Mid(st, 1, 1)) = 2 Then mojisu = mojisu + 2
    For i = 2 To l
        If LenB(Mid(st, i, 1)) = 2 Then
            mojisu = mojisu + 1
            If LenB(Mid(st, i - 1, 1)) = 1 Then mojisu = mojisu + 1
        Else
            If LenB(Mid(st, i - 1, 1)) = 2 Then mojisu = mojisu + 1
        End If
    Next
    
End Function

Sub findfromdatas()
Attribute findfromdatas.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim i As Integer, n As Integer, e As Boolean
    Dim gyosha As Integer, shina() As String, sh As Variant
    Dim ds As Variant
    Dim fl As Boolean
    
    Application.Calculation = xlCalculationManual
    
    e = False
    r = Selection.Row
    
    If Cells(r, Rgyosha) <> "" Then
        gyosha = Cells(r, Rgyosha)
    Else
        gyosha = Cells(r - 1, Rgyosha)
    End If

    tmp = ""
    For i = 0 To 2
     If Cells(Lcheck, Rshina + i) = True Then
      tmp = tmp & StrConv(UCase(Cells(r, Rshina + i)), vbNarrow)
     End If
    Next

    shina = Split(tmp, " ")
    
    
    i = 1
    n = 0
    While e = False
        If Int(Sheets("data").Cells(i, Dgyosha)) = Int(gyosha) Then
            fl = True
            For Each sh In shina
              If InStr(Sheets("data").Cells(i, Dshina), sh) = 0 Then fl = False
            Next
            If fl = True Then
                r1 = r + 3 + n
                Cells(r1, R678) = Sheets("data").Cells(i, D678)
                Cells(r1, Rgyosha) = Sheets("data").Cells(i, Dgyosha)
                Cells(r1, Rmnum) = Sheets("data").Cells(i, Dmnum)
                Cells(r1, Rfutan) = Sheets("data").Cells(i, Dfutan)
                Cells(r1, Rkanko) = Sheets("data").Cells(i, Dkanko)
                Cells(r1, Rkansht) = Sheets("data").Cells(i, Dkansht)
                Cells(r1, Rsuryo) = Sheets("data").Cells(i, Dsuryo)
                Cells(r1, Runit) = Sheets("data").Cells(i, Dunit)
                Cells(r1, Rtanka) = Sheets("data").Cells(i, Dtanka)
                Cells(r1, Rshina) = Sheets("data").Cells(i, Dshina)

                
                n = n + 1
            End If
        End If
        i = i + 1
        If n >= 100 Then e = True
        If Sheets("data").Cells(i, Dshina) = "" And Sheets("data").Cells(i, D678) = "" Then e = True

    Wend
    
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub bunkform()
Attribute bunkform.VB_ProcData.VB_Invoke_Func = "h\n14"
    ActiveSheet.CheckBoxes(1) = True
    ActiveSheet.CheckBoxes(2) = True
    ActiveSheet.CheckBoxes(3) = True
    
    UserForm1.TextBox1.Text = ""
    UserForm1.Show
    Call UserForm1.TextBox1.SetFocus
End Sub
