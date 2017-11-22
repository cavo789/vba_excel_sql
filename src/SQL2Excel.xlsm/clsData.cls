VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------------------------------------------------------------
'
' Author       : AVONTURE Christophe
'
' Aim          : Class Data connections
'                Allow to manipulate SQL Server data in Excel
'
' Written date : November 2017
'
' --------------------------------------------------------------------------------

Option Explicit
Option Compare Text
Option Base 0

Private cn As ADODB.Connection ' <== Requires Microsoft ActiveX Data Objects 2.8 Library
Private rs As ADODB.Recordset

Private Const cDSN As String = _
    "Driver={SQL Server};Server={0};Database={1};User Id={2};Password={3};"

Private p_ServerName As String
Private p_DBName As String
Private p_UserName As String
Private p_Password As String

Property Get ServerName() As String
    ServerName = p_ServerName
End Property

Property Let ServerName(ByVal sValue As String)
   p_ServerName = sValue
End Property

Property Get DatabaseName() As String
    DatabaseName = p_DBName
End Property

Property Let DatabaseName(ByVal sValue As String)
   p_DBName = sValue
End Property

Property Get UserName() As String
    UserName = p_UserName
End Property

Property Let UserName(ByVal sValue As String)
   p_UserName = sValue
End Property

Property Get Password() As String
    Password = p_Password
End Property

Property Let Password(ByVal sValue As String)
   p_Password = sValue
End Property

Property Get DSN() As String

Dim sDSN As String

    sDSN = StringFormat(cDSN, p_ServerName, p_DBName, p_UserName, p_Password)
    
    ' If no username was supplied, consider a trusted connection
    If (p_UserName = "") Then
        sDSN = sDSN & "Trusted_Connection=True;"
    End If
    
    DSN = sDSN
    
End Property

' --------------------------------------------------------------------------------
'
' Very quick control to check if the class has been correctly initialized
'
' --------------------------------------------------------------------------------

Property Get isReady(Optional ByVal sSQL As String = "") As Boolean

Dim bReady As Boolean

    bReady = False
    
    If (ServerName = "") Then
    
        MsgBox "ERROR - No server name has been specified", vbExclamation + vbOKOnly
        
    ElseIf (DatabaseName = "") Then
    
        MsgBox "ERROR - No database name has been specified", vbExclamation + vbOKOnly
        
    Else
    
        bReady = True
        
        If Not (IsMissing(sSQL)) Then
            If (Trim(sSQL) = "") Then
                bReady = False
                MsgBox "ERROR - Your SQL statement is empty", vbExclamation + vbOKOnly
            End If
        End If
        
    End If
    
    isReady = bReady
    
End Property

' --------------------------------------------------------------------------------
'
' Avoid string concatenations by using a syntax near of the
' .NET syntax : StringFormat
'
' http://www.self-access.com/cms/vb/vb-init/concatenation-vb/concatenation-abolir/
'
' Sample code:
' -----------
'
' Dim sTemp as string
' Dim sFirstName As string, sLastName As string, wAge as byte
'
'   sFirstName = "Christophe": sLastName = "Avonture": wAge = 20
'
'   Debug.Print CAVO_AddIn.cFunctions.StringFormat("Welcome {0} {1}, " & _
'       "it seems that you've {2} years. Really?", sFirstName, sLastName, wAge)
'
' --------------------------------------------------------------------------------

Private Function StringFormat(ByVal sString As String, ParamArray arrValues() As Variant) As String
 
Dim i As Byte, J As Byte, K As Byte
  
   J = LBound(arrValues)
   K = UBound(arrValues)
   
   For i = J To K
      sString = Replace(sString, "{" & i & "}", IIf(IsError(arrValues(i)), Null, arrValues(i)))
   Next
 
   StringFormat = sString
  
End Function

' --------------------------------------------------------------------------------
'
' Get a recordset from the database and output it into a sheet.
' This function make a copy not a link => there is no link
' with the database, no way to make a refresh.
'
' Advantage    : Fast
' Inconvenient : Don't keep any link with the DB, records are copied to Excel
'
' Sample code:
' -----------
'
' Dim cData As New clsData
' Dim rng As Range
'
'    cData.ServerName = "srvName"
'    cData.DatabaseName = "dbName"
'
'    ' When cData.UserName and cData.Password are not supplied
'    ' the connection will be made as "trusted" i.e. with the connected
'    ' user's credentials
'
'    Set rng = cData.SQL_CopyToSheet("SELECT TOP 5 * FROM tbl", ActiveSheet.Range("B2"))
'
' --------------------------------------------------------------------------------

Public Function SQL_CopyToSheet(ByVal sSQL As String, ByVal rngTargetCell As Range) As Range

Dim i As Long, J As Long
Dim rng As Range

    Application.EnableCancelKey = XlEnableCancelKey.xlDisabled

    If Not (Me.isReady(sSQL)) Then
        Exit Function
    End If

    If (rngTargetCell.Cells.Count > 1) Then
        MsgBox "ERROR - You should only specify the top-left cell where to copy " & _
           "the recordset (f.i. ActiveSheet.Range('A1'))", vbExclamation + vbOKOnly
        Set SQL_CopyToSheet = Nothing
        Exit Function
    End If

    Set cn = New ADODB.Connection
    
    If (cn.State <> 0) Then cn.Close
    cn.Open (Me.DSN())
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set rs = cn.Execute(sSQL)
    
    ' Export column's names
    J = rs.Fields.Count - 1
    
    For i = 0 To J
        rngTargetCell.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next
    
    Range(rngTargetCell.Cells(1, 1), rngTargetCell.Cells(1, J + 1)).Font.Bold = True
    
    ' And now export records
    rngTargetCell.Cells(2, 1).CopyFromRecordset rs
    
    rs.Close
    Set rs = Nothing
         
    cn.Close
    Set cn = Nothing
    
    ' Return the full range of the extracted informations (f.i $A$1:$J$5000)
    ' so the caller can work more easily with the returned range
    
    Set SQL_CopyToSheet = rngTargetCell.CurrentRegion

End Function

' --------------------------------------------------------------------------------
'
' Create a query table in a sheet : create the connection, the query table, give
' it a name and get data.
'
' Advantage    : * Keep the connection alive. The end user will be able to make a
'                  Data -> Refresh to obtain an update of the sheet
'                * If the user don't have access to the database, the records will
'                  well be visible but without any chance to refresh them
' Inconvenient : If the parameter bPersist is set to True, the connection string
'                will be in plain text in the file (=> avoid this if you're using
'                a login / password).
'
' Parameters
' ----------
'
'   sSQL       : Instruction to use (a valid SQL statement like
'                SELECT ... FROM ... or EXEC usp_xxxx)
'   sQueryName : Internal name that will be given to the querytable
'   rngTarget  : Destination of the returned recordset (f.i. Sheet1!$A$1)
'   bPersist   : If true, the connection string will be stored and, then, the
'                user will be able to make a refresh of the query
'
'                SECURIY CONCERN : IF USERNAME AND PASSWORD HAVE BEEN SUPPLIED,
'                THESE INFORMATIONS WILL BE SAVED IN CLEAR IN THE CONNECTION STRING !
'
' Sample code
' -----------
'
' Dim cData As New clsData
' Dim sSQL As String
'
'    cData.ServerName = "srvName"
'    cData.DatabaseName = "dbName"
'    cData.UserName = "JohnDoe"
'    cData.Password = "Password"
'
'    ' When cData.UserName and cData.Password are not supplied
'    ' the connection will be made as "trusted" i.e. with the connected
'    ' user's credentials
'
'    sSQL = "SELECT TOP 1 * FROM tbl"
'
'    Call cData.AddQueryTable(sSQL, "qryTest", ActiveCell, True)
'
' --------------------------------------------------------------------------------

Public Sub AddQueryTable(ByVal sSQL As String, ByVal sQueryName As String, _
    ByVal rngTarget As Range, Optional ByVal bPersist As Boolean = False, _
    Optional ByVal bAutoformat As Boolean = True)

Dim bEnableEvents As Boolean
Dim qt As QueryTable
Dim sh As Worksheet
Dim wb As Workbook
Dim wCount As Integer
   
    Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
    
    If Not (Me.isReady(sSQL)) Then
        Exit Sub
    End If
    
    bEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
    
    ' rngTarget is something like $A$1.
    ' Get Parent.Name to retrieve the name of the sheet (f.i. Sheet1)
    ' Get Parent.Parent.Name to retrieve the name of the workbook (f.i. Book1)
    
    Set wb = Application.Workbooks(rngTarget.Parent.Parent.Name)
    Set sh = wb.Worksheets(rngTarget.Parent.Name)
    
    ' rngTarget = Destination of the returned recordset (f.i. Sheet1!$A$1)
    Set qt = sh.QueryTables.Add(Connection:="OLEDB;" & Me.DSN(), Destination:=rngTarget)
    
    ' Name that will be given to the returned data (from row 1 till the last one)
    qt.Name = "_tblData"
    
    With qt
        .CommandText = sSQL
        .SavePassword = bPersist ' Should we store the password in the connection string ?
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .BackgroundQuery = False
    
        ' If an error occurs here when the SQL statement is well correct,
        ' copy/paste the SQL statement in SQL Server Management Studio
        ' when you've connect the server with the same credentials
        ' because it well possible that the SQL statement refers f.i. a table
        ' that is accessible to the dbo user but not to the application user
        ' used in this context.
        .Refresh
    
    End With
    
    ' Give names
    
    sh.Range(qt.Name).Rows(1).Select
    Call sh.Names.Add(qt.Name & "_Header", "=" & Selection.Address)
    
    ' Give a name on the first data row
    sh.Range(qt.Name).Rows(2).Select
    Call sh.Names.Add(qt.Name & "_FirstRow", "=" & Selection.EntireRow.Address)
        
    ' ----------------------------------
    ' Give a name to the connection; like f.i. "qryTest"
    On Error Resume Next
    wCount = wb.Connections.Count
    If ((Err.Number = 0) And (wCount > 0)) Then wb.Connections(wCount).Name = sQueryName
    On Error GoTo 0
        
    If bAutoformat Then
     
        ' Add AutoFilter and AutoFit columns
        On Error Resume Next
        sh.Range(qt.Name).Cells(1, 1).Select
        Selection.AutoFilter
        Selection.Columns.EntireColumn.AutoFit
        On Error GoTo 0
        
        ' For esthetics purposes, select the first data cell
        sh.Range(qt.Name).Cells(2, 1).Select
        
        ActiveWindow.FreezePanes = True
        
    End If
    
    ' Remove the QueryTable if not needed
    If (bPersist = False) Then
        On Error Resume Next
        qt.Delete
        On Error GoTo 0
    End If
    
    Application.EnableEvents = bEnableEvents
   
End Sub

' --------------------------------------------------------------------------------
'
' This function will call the AddQueryTable function of this class but first will
' create a new workbook, get the data and format the sheet (add a title, display the
' "Last extracted date" date/time in the report, add autofiltering, page setup and more.
'
' The obtained workbook will be ready to be sent to someone.
'
' Parameters
' ----------
'
'   sSQL         : Instruction to use (a valid SQL statement like
'                  SELECT ... FROM ... or EXEC usp_xxxx)
'   sReportTitle : Title for the sheet
'   bPersist     : If true, the connection string will be stored and, then, the
'                  user will be able to make a refresh of the query
'
'                  SECURIY CONCERN : IF USERNAME AND PASSWORD HAVE BEEN SUPPLIED,
'                  THESE INFORMATIONS WILL BE SAVED IN CLEAR IN THE CONNECTION STRING !
'
' Sample code
' -----------
'
' Dim cData As New clsData
' Dim sSQL As String
'
'    cData.ServerName = "srvName"
'    cData.DatabaseName = "dbName"
'    cData.UserName = "JohnDoe"
'    cData.Password = "Password"
'
'    sSQL = "SELECT TOP 1 * FROM tbl"
'
'    Call cData.RunSQLAndExportNewWorkbook(sSQL, "My Title", False)
'
' --------------------------------------------------------------------------------

Public Sub RunSQLAndExportNewWorkbook(ByVal sSQL As String, _
    ByVal sReportTitle As String, Optional bPersist As Boolean = True)

Dim sDefault As String
Dim sQueryName As String
Dim sTableName As String, sFileName As String
Dim sCommandText As String
Dim wb As Workbook
Dim bEvents As Boolean, bAlerts As Boolean
Dim wRow As Byte
Dim bProd As Boolean
Dim sh As Worksheet
   
    Application.EnableCancelKey = XlEnableCancelKey.xlDisabled

    If Not (Me.isReady(sSQL)) Then
        Exit Sub
    End If
    
    If ((Left(sSQL, 7) <> "SELECT ") And (Left(sSQL, 5) <> "EXEC ") And (Left(sSQL, 5) <> "WITH ")) Then
        MsgBox "ERROR - Only a SELECT statement can be fired; nothing else"
        Exit Sub
    End If
    
    If (sSQL <> vbNullString) And (sSQL <> sDefault) Then
        
        With Application
            bEvents = .EnableEvents
            bAlerts = .DisplayAlerts
            .EnableEvents = False
            .DisplayAlerts = False
        End With
    
        ' Query is being fired...
        
        Set wb = Workbooks.Add
        
        ' Keep only one sheet in our new workbook
        On Error Resume Next
        Do While wb.Worksheets.Count > 1
            wb.Worksheets(wb.Worksheets.Count).Delete
        Loop
        Err.Clear
        On Error GoTo 0
    
        Set sh = wb.ActiveSheet
        
        wRow = 2
        
        sh.Range(sh.Cells(wRow, 1), sh.Cells(wRow, 10)).Merge
        With sh.Cells(wRow, 1)
            .Value = sReportTitle
            .Font.Bold = True
            .Font.Size = 18
        End With
        
        wRow = wRow + 1
        sh.Range(sh.Cells(wRow, 1), sh.Cells(wRow, 10)).Merge
        With sh.Cells(wRow, 1)
            .Value = "Extraction date : " & Format(Now(), "dd/mm/yyyy hh:mm")
            .Font.Italic = True
            .Font.Size = 8
        End With
        
        Call Me.AddQueryTable(sSQL, "qry", sh.Cells(wRow + 2, 1), bPersist)
    
        ' Don't keep the connection : only data without possibility to make a refresh
        'Call cData.RemoveQueries(wb)
        
        ' No gridlines
        ActiveWindow.DisplayGridlines = False
        
        ' Freeze rows/columns
        ActiveWindow.FreezePanes = True
        sh.Cells(wRow + 3, 1).Select
        
        sh.Columns.AutoFit
        
        On Error Resume Next
        
        With sh.PageSetup
            .CenterFooter = "&""-,Bold""&14" & sReportTitle
            .RightFooter = "Page &P/&N"
            .Orientation = XlPageOrientation.xlLandscape
            .FitToPagesWide = 1
            .FitToPagesTall = 999
            .PrintTitleRows = "$1:$" & wRow + 2
            .Zoom = False
        End With
        
        Err.Clear
    
        On Error GoTo 0
        
        ' Get a temporay filename and remove the .tmp extension to let Excel to put the extension depending on the filetype (.xls or .xlsx)
        
        'sFileName = cFile.RemoveExtension(cFile.GetTempFileName)
        'wb.SaveAs sFileName
        
        Set sh = Nothing
        Set wb = Nothing
        
        With Application
            .EnableEvents = bEvents
            .DisplayAlerts = bAlerts
        End With
    
    End If ' If (sSQL <> vbNullString ) And (sSQL <> sDefault) Then
    
End Sub

