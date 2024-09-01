VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayPush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// 配列の末尾に追加する
Public Sub ArrayPush(ar As Variant, addValue As Variant)
    '// 引数が配列でない場合処理せず抜ける
    If IsArray(ar) = False Then
        Exit Sub
    End If
    
    '// 配列要素数を取得
    Dim iSize As Long
    
    If UBound(ar) = 0 And ar(0) = "" Then
        ar(0) = addValue
        Exit Sub
    End If
    
    '// 配列サイズを拡張後のサイズで取得
    iSize = UBound(ar) + 1
    
    '// 拡張
    ReDim Preserve ar(iSize)
    Dim i As Long         '// ループカウンタ
    
    '// オブジェクト型変数の場合
    If IsObject(ar(0)) = True Then
        '// 終端に現在ループ値を設定
        Set ar(iSize) = addValue
    '// プリミティブ型変数（IntegerやStringなど）の場合
    Else
        '// 終端に現在ループ値を設定
        ar(iSize) = addValue
    End If
    
End Sub
