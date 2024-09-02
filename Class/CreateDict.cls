VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' **********************************
' 辞書型作成(詳細) createNestedDictionary
' 1行目にKey設定
' 1列目にdictParentカラム名
' 2列目以降にdictChild設定
'{
'    "cell(2,1)": {
'        "cell(1,2)": "cell(2,2)",
'        "cell(1,3)": "cell(2,3)"
'    },
'    "cell(3,1)": {
'        "cell(1,2)": "cell(3,2)",
'        "cell(1,3)": "cell(3,3)"
'    },
'    "cell(4,1)": {
'        "cell(1,2)": "cell(4,2)",
'        "cell(1,3)": "cell(4,3)"
'    }
'}
' **********************************
Public Function createNestedDict(ByVal tblName As String) As Dictionary
    Dim dictChild As Object
    Dim dictParent As Object
    Dim columnFlg As Boolean
    
    Dim i As Integer
    Dim j As Integer
    
    i = 2
    Set dictParent = CreateObject("Scripting.Dictionary")
    Do
        j = 2
        Set dictChild = CreateObject("Scripting.Dictionary")
        Do While Worksheets(tblName).Cells(1, j) <> ""
            columnFlg = False
            dictChild(Worksheets(tblName).Cells(1, j).Value) = Worksheets(tblName).Cells(i, j).Value
            If Worksheets(tblName).Cells(i, j).Value <> "" Then
                columnFlg = True
            End If
            j = j + 1
        Loop
        If columnFlg = False Then Exit Do
        dictParent.Add (Worksheets(tblName).Cells(i, 1).Value), dictChild
        i = i + 1
    Loop
    
    Set createNestedDict = dictParent
End Function

' **********************************
' 辞書型作成(簡易) createCellMappingDict
' 1列目にkey設定
' 2列目にvalue設定
'{
'    "cell(1,1)":"cell(1,2)",
'    "cell(2,1)":"cell(2,2)",
'    "cell(3,1)":"cell(3,2)",
'    "cell(4,1)":"cell(4,2)",
'    "cell(5,1)":"cell(5,2)"
'}
' **********************************
Public Function createCellMappingDict(ByVal tblName As String) As Dictionary
    Dim dict As Object
    Dim columnFlg As Boolean
    
    Dim i As Integer
    Dim j As Integer
    
    i = 1
    Set dict = CreateObject("Scripting.Dictionary")
    
    Do While Worksheets(tblName).Cells(i, 1) <> ""
        dict(Worksheets(tblName).Cells(i, 1).Value) = Worksheets(tblName).Cells(i, 2).Value
        i = i + 1
    Loop

    Set createCellMappingDict = dict
End Function

' **********************************
' 配列作成 createCellComparisonList
' 1行目にkey設定
' 2行目以降にvalue設定
'[
'    {
'        "cell(1,1)":"cell(2,1)",
'        "cell(1,2)":"cell(2,2)",
'        "cell(1,3)":"cell(2,3)",
'        "cell(1,4)":"cell(2,4)",
'        "cell(1,5)":"cell(2,5)"
'    },
'    {
'        "cell(1,1)":"cell(3,1)",
'        "cell(1,2)":"cell(3,2)",
'        "cell(1,3)":"cell(3,3)",
'        "cell(1,4)":"cell(3,4)",
'        "cell(1,5)":"cell(3,5)"
'    }
']
'}
' **********************************

Public Function createCellComparisonList(ByVal tblName) As Variant
    Dim dict As Object
    Dim arr() As Variant
    Dim i As Integer
    Dim j As Integer
    Dim rowFlg As Boolean
    
    
    i = 2
    Do
        j = 1
        Set dict = CreateObject("Scripting.Dictionary")
        Do While Worksheets(tblName).Cells(1, j) <> ""
            rowFlg = False
            dict(Worksheets(tblName).Cells(1, j).Value) = Worksheets(tblName).Cells(i, j).Value
            If Worksheets(tblName).Cells(i, j).Value <> "" Then
                rowFlg = True
            End If
            j = j + 1
        Loop
        If rowFlg = False Then Exit Do
        ReDim Preserve arr(i - 2)
        Set arr(i - 2) = New Dictionary
        Set arr(i - 2) = dict
        i = i + 1
    Loop
    
    createCellComparisonList = arr

End Function
