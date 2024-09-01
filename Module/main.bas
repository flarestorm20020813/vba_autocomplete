Attribute VB_Name = "main"
Option Explicit
Public aryPush As ArrayPush
Public cmn As Common
Public dlg As FileDialog
Public jsn As JsonReader
Const playlists As String = "playlists"


Sub setup()
    Set aryPush = New ArrayPush
    Set cmn = New Common
    Set jsn = New JsonReader
End Sub


Sub pushButton()
    Dim playlistsdbJsonKeyArray() As Variant
    playlistsdbJsonKeyArray = Array("title", "_id")
    Dim trackidsdbJsonKeyArray() As Variant
    trackidsdbJsonKeyArray = Array("_id", "title")
    
    Dim playlistsdb As String
    Dim tracksdb As String
    
    playlistsdb = Worksheets("main").Cells(1, 2).Value
    tracksdb = Worksheets("main").Cells(2, 2).Value
    
    Call startJsonToSheet(playlistsdb, tracksdb, playlistsdbJsonKeyArray, trackidsdbJsonKeyArray)
End Sub


Public Sub startJsonToSheet(playlistsdb As String, tracksdb As String, playlistsdbJsonKeyArray As Variant, trackidsdbJsonKeyArray As Variant)
    '各種変数設定
    Dim ws As Worksheet, flag As Boolean
    Dim playlistsLineArray() As String
    Dim tracksLineArray() As String
    Dim i As Integer
    Dim elm As Object
    Dim json As String
    Dim item As Variant
    Dim j As Integer
    Dim k As Integer
    Dim a As String
    Dim sheetName As String
    
    '開始時セットアップ
    Call setup
    'Call ProcessControl.StartProcess
    
    'シートを削除(初期化)
    For Each ws In Worksheets
        If ws.Name <> "main" And ws.Name <> "song_list" Then Call cmn.delWorkSheet(ws.Name)
    Next ws
    
    'tracks.dbを1行ずつ格納する
    tracksLineArray = jsn.fileReadLineToArray(tracksdb)
    Call cmn.makeNewWorkSheet("tracksdb")
    
    
    For i = 0 To UBound(tracksLineArray)
        json = tracksLineArray(i)
        Set elm = JsonConverter.ParseJson(json)
        '情報をシートに
        For k = 0 To UBound(trackidsdbJsonKeyArray)
            If i = 0 Then
                Worksheets("tracksdb").Cells(1, k + 1).Value = trackidsdbJsonKeyArray(k)
            End If
            Worksheets("tracksdb").Cells(i + 2, k + 1).Value = CStr(elm(trackidsdbJsonKeyArray(k)))
        Next k
        Worksheets("tracksdb").Cells(i + 2, k + 1).Value = CStr(elm("file")("uri"))
    Next i
    
    
    'playlists.dbを1行ずつ格納する
    playlistsLineArray = jsn.fileReadLineToArray(playlistsdb)
    
    Call cmn.makeNewWorkSheet(playlists)
    For i = 0 To UBound(playlistsLineArray)
        j = 0
        json = playlistsLineArray(i)
        Set elm = JsonConverter.ParseJson(json)
        sheetName = elm("title")
        Call cmn.makeNewWorkSheet(sheetName)
        
        '情報をシートに
        For k = 0 To UBound(playlistsdbJsonKeyArray)
            If i = 0 Then
                Worksheets(playlists).Cells(1, k + 1).Value = playlistsdbJsonKeyArray(k)
            End If
            Worksheets(playlists).Cells(i + 2, k + 1).Value = CStr(elm(playlistsdbJsonKeyArray(k)))
        Next k
        
        'Worksheets(playlists).Cells(i + 2, 1).Value = sheetName
        

        For Each item In elm("_trackIds")
            j = j + 1
            Worksheets(sheetName).Cells(j, 1).Value = CStr(item)
            Worksheets(sheetName).Cells(j, 2).Formula = "=VLOOKUP($A" & CStr(j) & ",tracksdb!$A:$C,2,FALSE)"
            Worksheets(sheetName).Cells(j, 3).Formula = "=VLOOKUP($A" & CStr(j) & ",tracksdb!$A:$C,3,FALSE)"
            Worksheets(sheetName).Cells(j, 4).Formula = "=VLOOKUP($C" & CStr(j) & ",song_list!$A:$J,6,FALSE)"
            
        Next item
        
    Next i
    
    
    ThisWorkbook.sheets("main").Activate
    Call ProcessControl.EndProcess
End Sub



Public Sub sheetToJson()
    Call setup
    Dim playlistsdb As String
    Dim textlineArray() As String
    Dim newPlaylistIds() As String
    Dim i As Integer
    Dim j As Integer
    Dim json As String
    Dim elm As Object
    Dim sheetName As String
    Dim newTextlineArray() As String
    ReDim newTextlineArray(0)
    
    playlistsdb = "新しいテキスト ドキュメント.json"
    textlineArray = jsn.fileReadLineToArray(playlistsdb)
    
    
    Call cmn.makeNewWorkSheet("playlist")
    For i = 0 To UBound(textlineArray)
        j = 1
        ReDim newPlaylistIds(0)
        json = textlineArray(i)
        Set elm = JsonConverter.ParseJson(json)
        sheetName = elm("playlist")
        
        '各シートの値を格納
        Call cmn.makeNewWorkSheet(sheetName)
        ThisWorkbook.sheets(sheetName).Activate
        Do While Cells(j, 1) <> ""
            Call aryPush.ArrayPush(newPlaylistIds, Cells(j, 1).Value)
            j = j + 1
        Loop
        elm("playlistids") = newPlaylistIds
        
        Call aryPush.ArrayPush(newTextlineArray, JsonConverter.ConvertToJson(elm))
    Next i
    
    'テキストファイルに上書き
    Call jsn.textOutput(newTextlineArray)
    
    Shell "C:\windows\explorer.exe " & playlistsdb & "\", vbNormalFocus
End Sub




Public Sub ForNext2(ByVal cellrow As Long, cellcolumn As Long)
    '対象の列でなければ抜ける
    If cellcolumn <> 1 Then
        Exit Sub
    End If
    
    Call setup
    Dim searchValue As String
    searchValue = Cells(cellrow, cellcolumn).Value
    
    Dim dropdownList() As String
    Cells(cellrow, cellcolumn + 1).Clear
    
    Dim c As Long
    
    ReDim dropdownList(0)
    'dropdownList(0) = ""
    Dim i As Integer
    
    'リストにある各セルの値と比較し、一致すれば追加
    For i = 1 To 120
        Dim cellListValue As String
        cellListValue = Worksheets("リスト").Cells(i, 1).Value
        If cellListValue Like searchValue + "*" Then
             Call aryPush.ArrayPush(dropdownList, cellListValue)
        End If
    Next i
    
    
    If UBound(dropdownList) = 0 Then
        'ドロップダウンリストを削除
        With Cells(cellrow, cellcolumn + 1).Validation
            .Delete
        End With
        Cells(cellrow, cellcolumn + 1).Value = "×"
        Exit Sub
    ElseIf UBound(dropdownList) = 1 Then
        Cells(cellrow, cellcolumn + 1).Value = dropdownList(1)
    
    End If
    
    'ドロップダウンリストを作成
    With Cells(cellrow, cellcolumn + 1).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=Join(dropdownList, ",")
        .ShowError = False
    End With

End Sub
