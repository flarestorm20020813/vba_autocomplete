VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/***
' * jsonテキストから辞書型に変換する
' ***/
Public Function readJson(ByVal json As String) As Object


    readJson = JsonConverter.ParseJson(json)
    
End Function

'/***
' * ファイルを読み込み、1行ずつ配列にして返す
' ***/
Public Function fileReadLineToArray(ByVal fileName As String) As Variant
    Dim filePath As String
    Dim textLine As String
    
    Dim ary() As String
    ReDim ary(0)
    
    With CreateObject("ADODB.Stream")
        .Open
        .Charset = "UTF-8" 'BOMあり、BOMなし両対応
        .LineSeparator = -1 '-1=CRLF
        .LoadFromFile fileName
        
        '1行毎に処理
        Do Until .EOS
            textLine = .ReadText(-2) '1行取り出す
            Call aryPush.ArrayPush(ary, textLine)
        Loop
        .Close
    End With
    
    fileReadLineToArray = ary
End Function


'/***
' * テキストファイルを読み込みテキストにして返す
' ***/
Public Function readFile(ByVal fileName As String) As String
On Error GoTo Err
    
    Dim buf As String
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile fileName
        buf = .ReadText
        .Close
    End With
    
    readFile = buf
    
    Exit Function
'エラー処理
Err:
    readFile = ""
End Function

'テキストファイル出力
Public Sub textOutput(ByVal obj As Variant)
    Dim i As Integer
    
    'ファイルを新規モードで開く
    Open "新しいテキスト ドキュメント.json" For Output As #1
    
    '開いたファイルに書き込む
    For i = 0 To UBound(obj)
        Print #1, obj(i)
    Next
    
    '開いたファイルを閉じる
    Close #1

End Sub
