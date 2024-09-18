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
'    "cells(2,keyColumn)": {
'        "cells(1,1)": "cells(2,1)",
'        "cells(1,2)": "cells(2,2)",
'        "cells(1,3)": "cells(2,3)"
'    },
'    "cells(3,keyColumn)": {
'        "cells(1,1)": "cells(3,1)",
'        "cells(1,2)": "cells(3,2)",
'        "cells(1,3)": "cells(3,3)"
'    },
'    "cells(4,keyColumn)": {
'        "cells(1,1)": "cells(4,1)",
'        "cells(1,2)": "cells(4,2)",
'        "cells(1,3)": "cells(4,3)"
'    }
'}
' **********************************
Public Function createNestedDict(tblName As String, keyColumn As Integer) As Dictionary
    Dim dictChild As Object
    Dim dictParent As Object
    Dim columnFlg As Boolean
    
    Dim i As Integer
    Dim j As Integer
    
    Dim keyString As String
    
    i = 2
    Set dictParent = CreateObject("Scripting.Dictionary")
    Do
        j = 1
        Set dictChild = CreateObject("Scripting.Dictionary")
        
        Do While Worksheets(tblName).Cells(1, j) <> ""
            columnFlg = False
            dictChild(Worksheets(tblName).Cells(1, j).value) = Worksheets(tblName).Cells(i, j).value
            If Worksheets(tblName).Cells(i, j).value <> "" Then
                columnFlg = True
            End If
            j = j + 1
        Loop
        If columnFlg = False Then Exit Do
        
        keyString = Worksheets(tblName).Cells(i, keyColumn).value
        If dictParent.Exists(keyString) = False Then
            dictParent.Add (keyString), dictChild
        End If
        i = i + 1
    Loop
    
    Set createNestedDict = dictParent
End Function

' **********************************
' 辞書型作成(簡易) createCellMappingDict
' 1列目にkey設定
' 2列目にvalue設定
'{
'    "cells(1,1)":"cells(1,2)",
'    "cells(2,1)":"cells(2,2)",
'    "cells(3,1)":"cells(3,2)",
'    "cells(4,1)":"cells(4,2)",
'    "cells(5,1)":"cells(5,2)"
'}
' **********************************
Public Function createCellMappingDict(tblName As String) As Dictionary
    Dim dict As Object
    Dim columnFlg As Boolean
    
    Dim i As Integer
    Dim j As Integer
    
    i = 1
    Set dict = CreateObject("Scripting.Dictionary")
    
    Do While Worksheets(tblName).Cells(i, 1) <> ""
        dict(Worksheets(tblName).Cells(i, 1).value) = Worksheets(tblName).Cells(i, 2).value
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
'        "cells(1,1)":"cells(2,1)",
'        "cells(1,2)":"cells(2,2)",
'        "cells(1,3)":"cells(2,3)",
'        "cells(1,4)":"cells(2,4)",
'        "cells(1,5)":"cells(2,5)"
'    },
'    {
'        "cells(1,1)":"cells(3,1)",
'        "cells(1,2)":"cells(3,2)",
'        "cells(1,3)":"cells(3,3)",
'        "cells(1,4)":"cells(3,4)",
'        "cells(1,5)":"cells(3,5)"
'    }
']
' **********************************

Public Function createCellComparisonList(tblName As String) As Variant
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
            dict(Worksheets(tblName).Cells(1, j).value) = Worksheets(tblName).Cells(i, j).value
            If Worksheets(tblName).Cells(i, j).value <> "" Then
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

' **********************************
' グループ作成 createCellGroupList
' 1行目にkey設定
' 2行目以降にvalue設定
' 1行目の値をKeyとしてリストを作成する
'{
'    "Cells(2, keyColumn)": [
'        {
'            "Cells(1, 2)": "Cells(2, 2)",
'            "Cells(1, 3)": "Cells(2, 3)",
'            "Cells(1, 4)": "Cells(2, 4)",
'            "Cells(1, 5)": "Cells(2, 5)",
'            "Cells(1, 6)": "Cells(2, 6)"
'        },
'        {
'            "Cells(1, 2)": "Cells(3, 2)",
'            "Cells(1, 3)": "Cells(3, 3)",
'            "Cells(1, 4)": "Cells(3, 4)",
'            "Cells(1, 5)": "Cells(3, 5)",
'            "Cells(1, 6)": "Cells(3, 6)"
'        },
'        {
'            "Cells(1, 2)": "Cells(9, 2)",
'            "Cells(1, 3)": "Cells(9, 3)",
'            "Cells(1, 4)": "Cells(9, 4)",
'            "Cells(1, 5)": "Cells(9, 5)",
'            "Cells(1, 6)": "Cells(9, 6)"
'        },
'        {
'            "Cells(1, 2)": "Cells(7, 2)",
'            "Cells(1, 3)": "Cells(7, 3)",
'            "Cells(1, 4)": "Cells(7, 4)",
'            "Cells(1, 5)": "Cells(7, 5)",
'            "Cells(1, 6)": "Cells(7, 6)"
'        }
'    ],
'    "Cells(4, keyColumn)": [
'        {
'            "Cells(1, 2)": "Cells(4, 2)",
'            "Cells(1, 3)": "Cells(4, 3)",
'            "Cells(1, 4)": "Cells(4, 4)",
'            "Cells(1, 5)": "Cells(4, 5)",
'            "Cells(1, 6)": "Cells(4, 6)"
'        },
'        {
'            "Cells(1, 2)": "Cells(8, 2)",
'            "Cells(1, 3)": "Cells(8, 3)",
'            "Cells(1, 4)": "Cells(8, 4)",
'            "Cells(1, 5)": "Cells(8, 5)",
'            "Cells(1, 6)": "Cells(8, 6)"
'        },
'        {
'            "Cells(1, 2)": "Cells(10, 2)",
'            "Cells(1, 3)": "Cells(10, 3)",
'            "Cells(1, 4)": "Cells(10, 4)",
'            "Cells(1, 5)": "Cells(10, 5)",
'            "Cells(1, 6)": "Cells(10, 6)"
'        }
'    ],
'    "Cells(5, keyColumn)": [
'        {
'            "Cells(1, 2)": "Cells(5, 2)",
'            "Cells(1, 3)": "Cells(5, 3)",
'            "Cells(1, 4)": "Cells(5, 4)",
'            "Cells(1, 5)": "Cells(5, 5)",
'            "Cells(1, 6)": "Cells(5, 6)"
'        },
'        {
'            "Cells(1, 2)": "Cells(6, 2)",
'            "Cells(1, 3)": "Cells(6, 3)",
'            "Cells(1, 4)": "Cells(6, 4)",
'            "Cells(1, 5)": "Cells(6, 5)",
'            "Cells(1, 6)": "Cells(6, 6)"
'        }
'    ]
'}
' **********************************

Public Function createCellListDict(tblName As String, keyColumn As Integer) As Dictionary

    Dim dictChild As Object
    Dim dictParent As Object
    Dim columnFlg As Boolean
    
    Dim i As Integer
    Dim j As Integer
    
    Dim test As String
    
    Dim keyString As String
    
    i = 2
    Set dictParent = CreateObject("Scripting.Dictionary")
    
    Do While Worksheets(tblName).Cells(i, 1) <> ""
        keyString = Worksheets(tblName).Cells(i, keyColumn).value
        
        j = 1
        Set dictChild = CreateObject("Scripting.Dictionary")
        Do While Worksheets(tblName).Cells(1, j) <> ""
            columnFlg = False
            dictChild(Worksheets(tblName).Cells(1, j).value) = Worksheets(tblName).Cells(i, j).value
            
            j = j + 1
        Loop
        
        If dictParent.Exists(keyString) = False Then
            Dim newArr() As Variant
            ReDim Preserve newArr(0)
            Set newArr(0) = New Dictionary
            Set newArr(0) = dictChild
            
            dictParent.Add (keyString), newArr
        Else
            Dim arr() As Variant
            Dim iSize As Integer
            arr() = dictParent(keyString)
            iSize = UBound(dictParent(keyString)) + 1
            
            ReDim Preserve arr(iSize)
            Set arr(iSize) = New Dictionary
            Set arr(iSize) = dictChild
            
            dictParent(keyString) = arr()
        End If
        
        i = i + 1
    Loop
    
    Set createCellListDict = dictParent

End Function


