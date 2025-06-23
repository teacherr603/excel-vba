
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>家のメモ</title>
  <style>
    body {
      font-family: "Courier New", monospace;
      background-color: #f4f4f4;
      padding: 20px;
      line-height: 1.6;
    }
    .content {
      background: white;
      padding: 20px;
      border-radius: 8px;
      white-space: pre-wrap;
      box-shadow: 0 0 10px rgba(0,0,0,0.1);
    }
  </style>
</head>
<body>
  <h2></h2>
  <div class="content">
Sub ユーザー別_画面IDと時間の頻度集計()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook, ws As Worksheet
    Dim userDict As Object
    Dim userList As Variant
    Dim i As Long, lastRow As Long
    Dim dataArr As Variant
    Dim userID As String, screenID As String, timeStr As String
    Dim comboKey As String
    Dim freqDict As Object
    Dim savePath As String
    
    ' ★ データのあるフォルダのパス（最後に \ を付ける）
    folderPath = "C:\Your\Folder\Path\"  ' ← ここを実際のパスに変更してください

    ' ★ 対象のuseridリスト（任意の15個など）
    userList = Array("user1234", "user5678", "user9999") ' ← 実際のIDに差し替えてください

    ' ★ ユーザーごとの辞書を作成（Dictionary内にDictionaryを格納）
    Set userDict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(userList)
        userDict(userList(i)) = CreateObject("Scripting.Dictionary")
    Next

    ' ★ フォルダ内のすべてのExcelファイルをループ
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        On Error Resume Next
        Set wb = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        If Err.Number <> 0 Then
            Debug.Print "開けないファイル: " & fileName
            Err.Clear
            fileName = Dir
            GoTo NextFile
        End If
        On Error GoTo 0

        ' ★ 各ワークシートを処理
        For Each ws In wb.Worksheets
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow < 2 Then GoTo NextSheet

            ' ★ 範囲を一括で配列に読み込み（処理高速化）
            dataArr = ws.Range("A2:D" & lastRow).Value

            ' ★ 各行をループ
            For i = 1 To UBound(dataArr, 1)
                userID = Trim(CStr(dataArr(i, 1)))     ' A列：ユーザーID
                screenID = Trim(CStr(dataArr(i, 2)))   ' B列：画面ID
                timeStr = Trim(CStr(dataArr(i, 3)))    ' C列：時間（そのまま使う）

                ' ★ 対象ユーザーの場合のみ処理
                If userDict.Exists(userID) Then
                    comboKey = screenID & "|" & timeStr ' 複合キーでカウント
                    Set freqDict = userDict(userID)
                    If freqDict.Exists(comboKey) Then
                        freqDict(comboKey) = freqDict(comboKey) + 1
                    Else
                        freqDict(comboKey) = 1
                    End If
                End If
            Next i
NextSheet:
        Next ws
        wb.Close False
NextFile:
        fileName = Dir
    Loop

    ' ★ 結果をユーザーごとのExcelファイルに出力
    For Each userID In userDict.Keys
        Set freqDict = userDict(userID)
        If freqDict.Count > 0 Then
            Set wb = Workbooks.Add
            Set ws = wb.Sheets(1)
            ws.Range("A1:D1").Value = Array("ユーザーID", "画面ID", "時間", "頻度")

            i = 2
            For Each comboKey In freqDict.Keys
                ws.Cells(i, 1).Value = userID
                ws.Cells(i, 2).Value = Split(comboKey, "|")(0)
                ws.Cells(i, 3).Value = Split(comboKey, "|")(1)
                ws.Cells(i, 4).Value = freqDict(comboKey)
                i = i + 1
            Next

            savePath = folderPath & userID & "_結果.xlsx"
            Application.DisplayAlerts = False
            wb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
            wb.Close False
            Application.DisplayAlerts = True
        End If
    Next

    MsgBox "集計が完了しました。", vbInformation
End Sub

  </div>
</body>
</html>

