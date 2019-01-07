Attribute VB_Name = "Module1"
Option Explicit

Sub bookhikaku()
Dim wbk As Workbook: Set wbk = ThisWorkbook
Dim wbk1 As Worksheet: Set wbk1 = wbk.Worksheets(1) 'wbk1に転記先のシート名を代入
Dim wbk2 As Workbook
Dim wbk3 As Workbook
Dim wsn As String: wsn = "サーバリスト"

Dim WBK_Col As Long: WBK_Col = 1 '比較ブックに転機する最初のアドレス
Dim WBK_Row As Long: WBK_Row = 2 '比較ブックに転機する最初のアドレス

Dim D1, D2, D3, D4 As Long 'ループ用の変数
Dim OutCAL As Integer: OutCAL = 16 '比較表の右側のスタートカラム

Dim RefFr As Range: Set RefFr = wbk1.Range("C1")
Dim RefTo As Range: Set RefTo = wbk1.Range("S1")

'配列系の変数
Dim f_test1 As Variant
Dim l_test1 As Variant
Dim f_test2 As Variant
Dim l_test2 As Variant

Dim Judg_Aray1 As Variant
Dim Lost_Aray1 As Variant
Dim Add_Aray As Variant


D3 = 0
wbk1.Range(Cells(3, 1), Cells(Rows.Count, Columns.Count)).ClearContents
wbk1.Range(Cells(3, 1), Cells(Rows.Count, Columns.Count)).Interior.ColorIndex = 0



'参照元ファイルを開く
Dim OpenFileName As String
MsgBox "参照元のファイルを選択してください"
  OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
  If OpenFileName <> "False" Then
    Workbooks.Open OpenFileName
  Else
    MsgBox "キャンセルされました"
    Exit Sub
  End If

  Dim FileName: FileName = Dir(OpenFileName)
  Set wbk2 = Workbooks(FileName)
  
  With wbk2.Worksheets(wsn)
  If .AutoFilterMode Then             ''オートフィルタが設定されていたら
      If .AutoFilter.FilterMode Then  ''絞り込みがされていたら
             .Range("A1").AutoFilter     ''オートフィルタを解除する
      End If
  End If
  End With
  
RefFr.Value = wbk2.Name

  
'参照先ファイルを開く
MsgBox "参照先のファイルを選択してください"
  OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
  If OpenFileName <> "False" Then
    Workbooks.Open OpenFileName
  Else
    MsgBox "キャンセルされました"
    Exit Sub
  End If
  FileName = Dir(OpenFileName)
  Set wbk3 = Workbooks(FileName)
  
  With wbk3.Worksheets(wsn)
  If .AutoFilterMode Then             ''オートフィルタが設定されていたら
      If .AutoFilter.FilterMode Then  ''絞り込みがされていたら
             .Range("A1").AutoFilter     ''オートフィルタを解除する
      End If
  End If
  End With
    
RefTo.Value = wbk3.Name

If wbk2.Name = wbk3.Name Then
    MsgBox "同じファイルを選択しています", vbExclamation
    Exit Sub
End If
  
MsgBox "処理を開始します"


'excelfeezスタート
Call FreezeExcel
  
'比較元シートの最終行代入
Dim Copy_WBR As Integer: Copy_WBR = wbk2.Worksheets(wsn).Cells(Rows.Count, 12).End(xlUp).Row

'比較先シートの最終行代入
Dim Reve_WBR As Integer: Reve_WBR = wbk3.Worksheets(wsn).Cells(Rows.Count, 12).End(xlUp).Row

'IPアドレスの配列から、生きているものだけを抽出する。
f_test1 = Create_First_Aray(wbk2.Worksheets(wsn), Copy_WBR)
l_test1 = isLive(f_test1)

f_test2 = Create_First_Aray(wbk3.Worksheets(wsn), Reve_WBR)
l_test2 = isLive(f_test2)

Judg_Aray1 = JudgAndCreate_Aray(l_test1, l_test2)
'judg_arayで、比較して違ったものを合体させた配列ができた。

Lost_Aray1 = FoundLostIP(l_test1, l_test2)
Add_Aray = FoundLostIP(l_test2, l_test1)

'Judg_Aray1を転記する処理。
'Judg_Aray1で最初の要素がemptyだと処理を回避
If Not IsEmpty(Judg_Aray1(1, 1)) Then
    For D1 = 1 To UBound(Judg_Aray1, 1)
        For D2 = 1 To UBound(Judg_Aray1, 2)
            If D2 < 15 Then
                wbk1.Cells(WBK_Row + D1, WBK_Col + D2).Value = Judg_Aray1(D1, D2)
                If wbk1.Cells(WBK_Row + D1, WBK_Col + D2).Value <> Judg_Aray1(D1, D2 + 14) Then
                    wbk1.Cells(WBK_Row + D1, WBK_Col + D2).Interior.ColorIndex = 6
                
                End If
                    
            Else
                wbk1.Cells(WBK_Row + D1, WBK_Col + 1 + D2).Value = Judg_Aray1(D1, D2)
                If wbk1.Cells(WBK_Row + D1, WBK_Col + 1 + D2).Value <> Judg_Aray1(D1, D2 - 14) Then
                    wbk1.Cells(WBK_Row + D1, WBK_Col + 1 + D2).Interior.ColorIndex = 6
                    
                End If
                
            End If
    
        Next D2
    
    Next D1
Else
    D1 = 0
End If

If Not IsEmpty(Lost_Aray1(1, 1)) Then
    For D3 = 1 To UBound(Lost_Aray1, 1)
        wbk1.Cells(WBK_Row + D1 + D3, WBK_Col).Value = "削除"
        wbk1.Cells(WBK_Row + D1 + D3, WBK_Col).Interior.ColorIndex = 3
        
        For D4 = 1 To UBound(Lost_Aray1, 2)
            wbk1.Cells(WBK_Row + D1 + D3, WBK_Col + D4).Value = Lost_Aray1(D3, D4)
        
        Next D4
    
    Next D3
End If

If Not IsEmpty(Add_Aray(1, 1)) Then
    For D3 = 1 To UBound(Add_Aray, 1)
        wbk1.Cells(WBK_Row + D1 + D3, OutCAL).Value = "追加"
        wbk1.Cells(WBK_Row + D1 + D3, OutCAL).Interior.ColorIndex = 6
        
        For D4 = 1 To UBound(Add_Aray, 2)
            wbk1.Cells(WBK_Row + D1 + D3, OutCAL + D4).Value = Add_Aray(D3, D4)
        
        Next D4
    
    Next D3
End If


'開いているブックを閉じる処理
wbk2.Close
wbk3.Close

MsgBox "処理終了"

Call MeltExcel

End Sub



'渡された二つの配列の中で、見つからなかったIPの配列を渡す。
Function FoundLostIP(Aray1 As Variant, Aray2 As Variant) As Variant
Dim FoundAray() As Variant
Dim i As Integer
Dim k As Integer
Dim m As Integer
Dim F_Found As Boolean
Dim Lost_Cnt As Integer
Dim Lost_Ins As Integer
Lost_Cnt = 0

F_Found = False

'一致しないIPの要素数を調べる。
For i = 1 To UBound(Aray1, 1)
    If Aray1(i, 11) Like "*.*" And Not IsEmpty(Aray1(i, 11)) Then
        For k = 1 To UBound(Aray2, 1)
            If Aray1(i, 11) = Aray2(k, 11) Then
                Exit For
                
            End If
            
            If k = UBound(Aray2, 1) Then
                Lost_Cnt = Lost_Cnt + 1
            
            End If
        
        Next k
            
    End If
 
Next i

If Lost_Cnt = 0 Then
    ReDim FoundAray(1 To 1, 1 To 14)
    FoundLostIP = FoundAray
    Exit Function
    
End If

'一致しないIPの要素数を二次元配列で宣言する。
ReDim FoundAray(1 To Lost_Cnt, 1 To 14)
F_Found = False

Lost_Ins = 1
For i = 1 To UBound(Aray1, 1)

    If Aray1(i, 11) Like "*.*" And Not IsEmpty(Aray1(i, 11)) Then
        For k = 1 To UBound(Aray2, 1)
            If Aray1(i, 11) = Aray2(k, 11) Then
                Exit For
                
            End If
            
            If k = UBound(Aray2, 1) Then
                For m = 1 To 14
                    FoundAray(Lost_Ins, m) = Aray1(i, m + 10)
        
                Next m
                Lost_Ins = Lost_Ins + 1
                
            End If
        
        Next k
        
    End If
 
Next i

FoundLostIP = FoundAray

End Function


'渡された二つの配列の中で、一致しないものを抽出した配列を作る関数
Function JudgAndCreate_Aray(Aray1 As Variant, Aray2 As Variant) As Variant
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim n As Integer

Dim Judg As Boolean
Dim Judg_Count As Integer
Dim C_Aray1() As Variant

Judg = False
Judg_Count = 0
Dim Aray_Count As Integer
Aray_Count = 1

'要素が変化しているIPアドレスの要素数を調べる。
For i = 1 To UBound(Aray1, 1)
    If Aray1(i, 11) Like "*.*" And Not IsEmpty(Aray1(i, 11)) Then
       For k = 1 To UBound(Aray2, 1)
            If Aray1(i, 11) = Aray2(k, 11) Then
                For j = 1 To (UBound(Aray1, 2) - 11)
                    If Aray1(i, 11 + j) <> Aray2(k, 11 + j) Then
                        Judg_Count = Judg_Count + 1
                        Exit For
                
                    End If
                
                Next j
            
                Exit For
            End If 'Aray1 = Aray2
       
       Next k
    End If

Next i

If Judg_Count = 0 Then
    ReDim C_aray(1 To 1, 1 To 28)
    JudgAndCreate_Aray = C_aray
    Exit Function
End If

'要素に変化があったIPの数を二次元配列に格納するための宣言を行う。
ReDim C_aray(1 To Judg_Count, 1 To 28)

'C_Arayには比較元の配列と比較先の配列が両方入っている。
'二次配列の15番で折り返し（14番までが参照元の配列）

For i = 1 To UBound(Aray1, 1)
    If Aray1(i, 11) Like "*.*" And Not IsEmpty(Aray1(i, 11)) Then
        For k = 1 To UBound(Aray2, 1)
            If Aray1(i, 11) = Aray2(k, 11) Then
                For j = 1 To (UBound(Aray1, 2) - 11)
                    If Aray1(i, 11 + j) <> Aray2(k, 11 + j) Then
                        For m = 1 To (UBound(C_aray, 2) / 2)
'ここでばぐった
                            C_aray(Aray_Count, m) = Aray1(i, 10 + m)
                            C_aray(Aray_Count, m + 14) = Aray2(k, 10 + m)
                    
                        Next m
                        
                        Aray_Count = Aray_Count + 1
                        Exit For
            
                    End If
                
                Next j
            
                Exit For
        
            End If
    
        Next k
    
    End If

Next i

JudgAndCreate_Aray = C_aray

End Function


'最終行の番号を入れると、B列からY列までの値を配列にして返す
Function Create_First_Aray(WS As Worksheet, LastRow As Integer) As Variant
'B列からY列まで値をとる
'B列は2、Y列は25
'行は4からスタート
Dim Parent_Range As Range
Set Parent_Range = WS.Range(WS.Cells(4, 2), WS.Cells(LastRow, 25))
Create_First_Aray = Parent_Range

End Function

'入ってきた配列から生きているものだけを抽出する。
'配列の１列目に「予約」の文字が入っている配列を消して返す。
Function isLive(Aray As Variant) As Variant
Dim Live_Cnt As Integer
Dim Live_Aray As Variant
'二次元配列の要素数入れる時に、Uboundが入らなかったので、下の変数でUboundを代入。
Dim Aray_Clm As Integer: Aray_Clm = UBound(Aray, 2)
Dim i, j As Integer
Dim NotIsLive_Cnt As Integer

Live_Cnt = 0
For i = 1 To UBound(Aray, 1)
    If Not Aray(i, 1) Like "*予約*" Then
        Live_Cnt = Live_Cnt + 1
    End If
Next i

ReDim Live_Aray(1 To Live_Cnt, 1 To Aray_Clm)

NotIsLive_Cnt = 1
For i = 1 To UBound(Aray, 1)
    If Not Aray(i, 1) Like "*予約*" Then
        For j = 1 To Aray_Clm
            Live_Aray(NotIsLive_Cnt, j) = Aray(i, j)
        Next j
        NotIsLive_Cnt = NotIsLive_Cnt + 1
    Else

    End If
Next i

isLive = Live_Aray

End Function


Private Function FreezeExcel()
    With Application
        '.Visible = False '全体の表示を停止
        .DisplayAlerts = False 'アラートの表示を停止
        .StatusBar = False 'ステータスバーの表示更新を停止
        .ScreenUpdating = False 'スクリーンの描画を停止
        .EnableEvents = False 'イベントを一時停止
        .Calculation = xlManual '計算を手動モードにする
    End With
End Function
   

Private Function MeltExcel()
    With Application
        .Calculation = xlAutomatic '計算を自動モードに戻す
        .EnableEvents = True 'イベントを再開
        .ScreenUpdating = True 'スクリーンの描画を再開
        .StatusBar = True 'ステータスバーの表示を再開
        .DisplayAlerts = True 'アラートの表示を再開
        '.Visible = True '全体の表示を回復
    End With
End Function
