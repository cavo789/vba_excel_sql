Attribute VB_Name = "test"
' --------------------------------------------------------------------------------
'
' Author       : AVONTURE Christophe
'
' Aim          : Test for clsData
'
' Written date : November 2017
'
' --------------------------------------------------------------------------------

Option Explicit
Option Base 0
Option Compare Text

Private Const cServerName = ""   ' <-- You'll need to mention here your server name
Private Const cDBName = ""       ' <-- You'll need to mention here your database name
Private Const cSQLStatement = "" ' <-- You'll need to mention here a valid SQL statement (SELECT ...)

Dim cData As New clsData

Sub CopyToSheet()

Dim rng As Range

    cData.ServerName = cServerName
    cData.DatabaseName = cDBName

    Set rng = cData.SQL_CopyToSheet(cSQLStatement, ActiveSheet.Range("A1"))
    
End Sub

Sub AddQueryTable()

    cData.ServerName = cServerName
    cData.DatabaseName = cDBName
    
    Call cData.AddQueryTable(cSQLStatement, "qryTest", ActiveSheet.Range("A1"), True)

End Sub

Sub RunSQLAndExportNewWorkbook()

    cData.ServerName = cServerName
    cData.DatabaseName = cDBName
        
    Call cData.RunSQLAndExportNewWorkbook(cSQLStatement, "My title", True)

End Sub
