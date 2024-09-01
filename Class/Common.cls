VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/***********
' * 画面上にメッセージを表示する
' ***********/
Public Sub myMsgBox6(Word As String, T As Single)
    Dim L As Shape    '←表示する図形
    Dim startT As Single    '←表示開始時の時刻
    Dim X As Double    '←図形の左上角の位置（X方向）
    Dim Y As Double    '←図形の左上角の位置（Y方向）
    With ActiveWindow.VisibleRange
        X = .Left + .Width / 2
        Y = .Top + .Height / 3
    End With
    Set L = ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, X, Y, 0, 0)
    L.TextFrame2.TextRange.Characters.Text = Word
    L.TextFrame2.WordWrap = msoFalse
    L.Fill.BackColor.RGB = RGB(0, 255, 255)
    L.Line.Visible = msoTrue
    L.Shadow.Type = msoShadow21
    startT = Timer()
    Do While (startT + T) > Timer()
        DoEvents: DoEvents
    Loop
    On Error Resume Next
        L.Delete
    On Error GoTo 0
    Set L = Nothing
End Sub

'/***********
' * シート削除
' ***********/
Public Sub delWorkSheet(sheetName As String)
    'シートがあるか確認し、無いなら関数を終了する
    Dim ws As Worksheet, flag As Boolean
    For Each ws In Worksheets
        If ws.Name = sheetName Then flag = True
    Next ws
    If flag = False Then
        Exit Sub
    End If
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
End Sub

'/***********
' * シート作成
' ***********/
Public Sub makeNewWorkSheet(sheetName As String)
    'シートがあるか確認し、あるなら関数を終了する
    Dim ws As Worksheet, flag As Boolean
    For Each ws In Worksheets
        If ws.Name = sheetName Then flag = True
    Next ws
    If flag = True Then
        Exit Sub
    End If
    
    Worksheets().Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = sheetName
End Sub
