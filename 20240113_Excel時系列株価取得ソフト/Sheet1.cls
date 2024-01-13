VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdKabukaGet_Click()
    On Error GoTo ErrorTrap
    Dim wRow        As Long
    Dim wNo         As Long
   
    '更新確認メッセージ
    If MsgBox("株価を取得します。よろしいですか？", vbQuestion + vbOKCancel, "確認") <> vbOK Then
        GoTo ExitTrap
    End If
    
    'マウスポインタを砂時計にする
    Application.Cursor = xlWait

    'A列の6行目から最終行まで処理を繰り返す
    wNo = 1
    For wRow = 6 To Range("A" & Rows.Count).End(xlUp).Row
        Me.Activate
        Me.Cells(wRow, 1).Select
        If (Me.Cells(wRow, 1) >= 1000 And Me.Cells(wRow, 1) <= 9999) _
         Or Me.Cells(wRow, 1) = 998405 Or Me.Cells(wRow, 1) = 998407 Then
            Call GetJikeretu(Cells(wRow, 1), wRow)
        End If
        wNo = wNo + 1
    Next
ExitTrap:
    'マウスポインタを通常に戻す
    Application.Cursor = xlDefault
    Exit Sub
ErrorTrap:
    'エラー処理
    MsgBox "cmdKabukaGet_Click Error!" & Err.Number & ":" & Err.Description, vbExclamation + vbOKOnly, "Error!!"
    Resume ExitTrap
End Sub

Private Sub GetJikeretu(prmCode, prmListRow)
    On Error GoTo ErrorTrap
    Dim oHttp       As Object
    Dim strURL      As String
    Dim strText     As String
    Dim arrData()   As String
    Dim arrLine()   As String
    Dim wIdx1       As Long
    Dim wIdx2       As Long
    Dim wRow        As Long
    Dim wClm        As Long
    Dim wMaxRow     As Long
    Dim wStrno      As Long
    Dim wEndno      As Long
    Dim wSheetName  As String
    Dim wSheetCnt   As Long
    Dim wPage       As Long
    
    wRow = 3
    wPage = 1
        
    wSheetName = "s" & prmCode
    If chkSheetUmu(wSheetName) = True Then
        Worksheets(wSheetName).Range("A:G").ClearContents   'シートクリア
    End If
    
    'Yahooファイナンス「時系列データ 」ページのHTMLソースを取得
    strURL = "https://finance.yahoo.co.jp/quote/" & prmCode & "/history?"
    strURL = strURL & "from=" & Format(Me.Range("開始日"), "yyyymmdd")
    strURL = strURL & "&to=" & Format(Me.Range("終了日"), "yyyymmdd")
    Select Case Range("期間単位")
        Case "週間"
            strURL = strURL & "&timeFrame=w"
        Case "月間"
            strURL = strURL & "&timeFrame=m"
        Case Else
            strURL = strURL & "&timeFrame=d"
    End Select
    strURL = strURL & "&page="

UrlGetTrap:
    'オブジェクト変数に参照セットする
    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    
    With oHttp
        'URL読み込み
        .Open "GET", strURL & wPage, False
        .Send
        
        If (.Status < 200 Or .Status >= 300) Then
            MsgBox "コード:" & prmCode & "のURL読み込みに失敗しました", vbExclamation + vbOKOnly, "Error!"
            GoTo ExitTrap
        End If
        
        If InStr(1, .responseText, "時系列情報がありません") > 0 Then
            GoTo ExitTrap
        End If

        If InStr(1, .responseText, "終値") = 0 Then
            MsgBox "コード:" & prmCode & "の時系列データが見つかりません", vbInformation + vbOKOnly
            GoTo ExitTrap
        End If
        
        If wPage = 1 Then
            If chkSheetUmu(wSheetName) = False Then
            'シート作成
                wSheetCnt = Worksheets.Count
                Worksheets.Add after:=Worksheets(wSheetCnt)
                ActiveSheet.Name = wSheetName
                '列幅設定
                With ActiveSheet
                    .Columns(1).ColumnWidth = 14
                    .Columns(2).ColumnWidth = 10
                    .Columns(3).ColumnWidth = 10
                    .Columns(4).ColumnWidth = 10
                    .Columns(5).ColumnWidth = 10
                    .Columns(6).ColumnWidth = 12
                    .Columns(7).ColumnWidth = 10
                End With
            End If
            Me.Hyperlinks.Add Anchor:=Me.Cells(prmListRow, 1), Address:="", SubAddress:="'" & wSheetName & "'!A1"
            Worksheets(wSheetName).Cells(1, 1) = prmCode
            
            'HTMLソースから[銘柄名]取り出し
            If prmCode = 998405 Or prmCode = 998407 Then
                strText = GetText(.responseText, "<title>", "：")
            Else
                strText = GetText(.responseText, "<title>", "【")
            End If
            Me.Cells(prmListRow, 2) = strText
            Worksheets(wSheetName).Cells(1, 2) = strText
            Worksheets(wSheetName).Hyperlinks.Add Worksheets(wSheetName).Cells(1, 2), strURL & wPage    'YahooファイナンスWebページをハイパーリンク
               
            '列見出し
            Worksheets(wSheetName).Cells(2, 1) = "日付"
            Worksheets(wSheetName).Cells(2, 2) = "始値"
            Worksheets(wSheetName).Cells(2, 3) = "高値"
            Worksheets(wSheetName).Cells(2, 4) = "安値"
            Worksheets(wSheetName).Cells(2, 5) = "終値"
            If prmCode <= 9999 Then
                Worksheets(wSheetName).Cells(2, 6) = "出来高"
                Worksheets(wSheetName).Cells(2, 7) = "調整後終値"
            End If
        End If
           
        'HTMLソースから時系列表データ切り出し
        If prmCode <= 9999 Then
            strText = GetText(.responseText, "histories" & Chr(34) & ":[", "paging")
            strText = GetText(strText, "histories" & Chr(34) & ":[", "}],")
        Else
            strText = GetText(.responseText, "mainDomesticIndexHistory", "paging")
            strText = GetText(strText, "histories" & Chr(34) & ":[", "}],")
        End If
        
        '1行区切りごとに配列セット
        arrData = Split(strText, Chr(34) & "},{", , vbTextCompare)
        
        For wIdx1 = LBound(arrData) To UBound(arrData)
            '1行分のデータを列ごとに配列セット
            arrLine = Split(arrData(wIdx1), Chr(34) & "," & Chr(34), , vbTextCompare)
            wClm = 1
            For wIdx2 = LBound(arrLine) To UBound(arrLine)
                strText = arrLine(wIdx2)
                
                '1項目分のデータから改行(Chr(10))を切り取る
                Do Until InStr(1, strText, Chr(10)) = 0
                    wStrno = InStr(1, strText, Chr(10))
                    strText = Mid(strText, 1, wStrno - 1) & Mid(strText, wStrno + 1, Len(strText) - wStrno)
                Loop
                
                '1項目分のデータから値のみ切り取る
                wStrno = InStr(1, strText, ":")
                strText = Mid(strText, wStrno + 2, Len(strText) - wStrno + 2)
                
                'セルへ値セット
                strText = Replace(strText, Chr(34), "")
                strText = Replace(strText, "}", "")
                
                If prmCode <= 9999 Then
                    Select Case wClm
                        Case 1: Worksheets(wSheetName).Cells(wRow, 2) = strText
                        Case 2: Worksheets(wSheetName).Cells(wRow, 3) = strText
                        Case 3: Worksheets(wSheetName).Cells(wRow, 4) = strText
                        Case 4: Worksheets(wSheetName).Cells(wRow, 5) = strText
                        Case 5: Worksheets(wSheetName).Cells(wRow, 6) = strText
                        Case 6: Worksheets(wSheetName).Cells(wRow, 7) = strText
                        Case 7
                            If IsDate(strText) Then
                                Worksheets(wSheetName).Cells(wRow, 1) = CDate(strText)
                            Else
                                Worksheets(wSheetName).Cells(wRow, 1) = strText
                            End If
                    End Select
                    If wClm >= 7 Then
                        Exit For
                    End If
                Else
                    Select Case wClm
                        Case 1
                            If IsDate(strText) Then
                                Worksheets(wSheetName).Cells(wRow, 1) = CDate(strText)
                            Else
                                Worksheets(wSheetName).Cells(wRow, 1) = strText
                            End If
                        Case 2: Worksheets(wSheetName).Cells(wRow, 2) = strText
                        Case 3: Worksheets(wSheetName).Cells(wRow, 3) = strText
                        Case 4: Worksheets(wSheetName).Cells(wRow, 4) = strText
                        Case 5: Worksheets(wSheetName).Cells(wRow, 5) = strText
                    End Select
                    If wClm >= 5 Then
                        Exit For
                    End If
                End If
                wClm = wClm + 1
            Next
            wRow = wRow + 1
        Next
        
        strText = .responseText
    End With
    
    '次ページのHTMLソースを取得
    wPage = wPage + 1
    GoTo UrlGetTrap
    
ExitTrap:
    '日付順に並び替え
    If Range("順序") = "日付昇順" Then
        If chkSheetUmu(wSheetName) = True Then
            Worksheets(wSheetName).Select
            Worksheets(wSheetName).Cells(3, 1).Select
            Selection.Sort Key1:=Worksheets(wSheetName).Cells(3, 1), Order1:=xlAscending, Header:=xlGuess, _
                OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin
        End If
    End If
    
    Me.Activate
    
    'オブジェクト変数を解放する
    Set oHttp = Nothing
    
    Exit Sub
ErrorTrap:
    'エラー処理
    MsgBox "cmdKabukaGet_Click Error!" & Err.Number & ":" & Err.Description, vbExclamation + vbOKOnly, "Error!!"
    Resume ExitTrap
End Sub

Public Function GetText(prmAllText As String, prmStrText, prmEndText)
    '全体文字列(prmAllText)の中から開始文字列(prmStrText)～終了文字列(prmEndText)までの間の文字を取得する
    Dim wStrno      As Long
    Dim wEndno      As Long
    
    wStrno = InStr(1, prmAllText, prmStrText) + Len(prmStrText)     '開始文字列の次の文字位置を取得する
    wEndno = InStr(wStrno, prmAllText, prmEndText)                  '終了文字列の位置を取得する
    GetText = Mid(prmAllText, wStrno, wEndno - wStrno)              '開始文字列～終了文字列までの間の文字を取得する
End Function

Function chkSheetUmu(prmSheetName) As Boolean
    'シートの存在有無チェック
    Dim wSheet As Object
    
    chkSheetUmu = False
    
    For Each wSheet In Sheets
        If LCase(prmSheetName) = LCase(wSheet.Name) Then
            chkSheetUmu = True
            Exit For
        End If
    Next
End Function

