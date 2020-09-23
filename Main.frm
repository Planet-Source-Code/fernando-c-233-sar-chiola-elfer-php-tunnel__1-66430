VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Main 
   Caption         =   "ElFerPHPTunnel"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Log 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   6960
      Width           =   12015
   End
   Begin VB.CommandButton cmdCreateTable 
      Caption         =   "Create Table"
      Height          =   300
      Left            =   10920
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdTableInfo 
      Caption         =   "TableInfo"
      Height          =   300
      Left            =   9600
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdShowTables 
      Caption         =   "ShowTables"
      Height          =   300
      Left            =   8280
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9763
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      Height          =   300
      Left            =   10920
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11640
      Top             =   960
   End
   Begin VB.TextBox Port 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Text            =   "80"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox QueryFilePath 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Text            =   "/wikilyric/query.php"
      Top             =   120
      Width           =   5295
   End
   Begin VB.TextBox Host 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "usuarios.lycos.es"
      Top             =   120
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   11640
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtQuery 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Main.frx":0000
      Top             =   480
      Width           =   12015
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txtResult As String
Private Sub Form_Load()
    MsgBox "Please VOTE, and dont drop the tables"
    txtQuery_KeyPress (0)
    txtQuery.SelStart = Len(txtQuery)
End Sub
Private Sub cmdCreateTable_Click()
    txtQuery = "CREATE TABLE " & InputBox("(Case sensitive) Table Name: ", "Create Table") & _
    "(" & InputBox("FieldName1 DataType1, FieldName2 DataType2, ...: ", "Fields") & ")"
    If txtQuery <> "CREATE TABLE " Then cmdQuery_Click
End Sub
Private Sub cmdQuery_Click()
    txtResult = ""
    Log = ""
    Timer.Enabled = True
    Winsock.Close
    Winsock.Connect Host, Port
End Sub
Private Sub cmdShowTables_Click()
    txtQuery = "SHOW TABLES"
    cmdQuery_Click
End Sub
Private Sub cmdTableInfo_Click()
    txtQuery = "DESCRIBE " & InputBox("(Case sensitive) Table Name: ", "Table Info")
    If txtQuery <> "DESCRIBE " Then cmdQuery_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Timer.Enabled = False
    txtResult = ""
    Log = ""
    Winsock.Close
    Unload Me
    DoEvents
    End
End Sub

Private Sub Timer_Timer()
Timer.Enabled = False
    Dim Result As String
    Dim Start As Long
    Start = 1
    Result = GetChunk(txtResult, 1, "<table>", "</table>")
    If Result <> "" Then
        txtResult = Result
        Rows = CountChunks(Result, 1, vbTab & "<tr>", "</tr>")
        Cols = CountChunks(Result, 1, vbTab & "<td>", "</td>") / Rows
        ReDim ArrayResult(Rows, Cols) As String
        For x = 0 To Rows - 1
            For y = 0 To Cols - 1
                ArrayResult(x, y) = GetChunk(txtResult, Start, "<td>", "</td>")
                ArrayResult(x, y) = Replace(ArrayResult(x, y), "<td>", "")
                ArrayResult(x, y) = Replace(ArrayResult(x, y), "</td>", "")
            Next
        Next
        txtResult = ""
        ArrayToMSFlexGrid ArrayResult, MSFlexGrid
    Else
        Timer.Enabled = True
        Result = GetChunk(txtResult, 1, "[ Error: ", " ]")
        If Result <> "" Then MsgBox Result, vbInformation, "Error": Timer.Enabled = False: txtResult = ""
    End If
End Sub
Private Function GetChunk(HtmlStr As String, Init As Long, First As String, Last As String) As String
    FirstPos = InStr(Init, HtmlStr, First)
    LastPos = InStr(FirstPos + Len(First), HtmlStr, Last)
    If FirstPos > 0 And LastPos > 0 Then
        GetChunk = Mid$(HtmlStr, FirstPos, LastPos + Len(Last) - FirstPos)
        Init = LastPos + Len(Last)
    End If
End Function
Private Function CountChunks(HtmlStr As String, Init As Long, First As String, Last As String) As Integer
    Do: DoEvents
        FirstPos = InStr(Init, HtmlStr, First)
        LastPos = InStr(FirstPos + Len(First), HtmlStr, Last)
        If FirstPos > 0 And LastPos > 0 Then
            CountChunks = CountChunks + 1
            Init = LastPos + Len(Last)
        Else
            Exit Do
        End If
    Loop
End Function
Private Sub ArrayToMSFlexGrid(Arr As Variant, Grid As MSFlexGrid)
    Grid.Rows = UBound(Arr, 1)
    Grid.Cols = UBound(Arr, 2)
    For x = 0 To Grid.Rows - 1
        For y = 0 To Grid.Cols - 1
            Grid.TextMatrix(x, y) = Arr(x, y)
        Next
    Next
    Grid.Col = 0: Grid.ColSel = Grid.Cols - 1
End Sub
Private Sub txtQuery_KeyPress(KeyAscii As Integer)
If txtQuery.SelLength > 0 Then Exit Sub
WordsArray = Split(txtQuery, " ")
S = txtQuery.SelStart
For x = 0 To UBound(WordsArray)
    If UCase(WordsArray(x)) = "SELECT" Then txtQuery = Replace(txtQuery, WordsArray(x), "SELECT")
    If UCase(WordsArray(x)) = "FROM" Then txtQuery = Replace(txtQuery, WordsArray(x), "FROM")
    If UCase(WordsArray(x)) = "WHERE" Then txtQuery = Replace(txtQuery, WordsArray(x), "WHERE")
    If UCase(WordsArray(x)) = "UPDATE" Then txtQuery = Replace(txtQuery, WordsArray(x), "UPDATE")
    If UCase(WordsArray(x)) = "INSERT" Then txtQuery = Replace(txtQuery, WordsArray(x), "INSERT")
    If UCase(WordsArray(x)) = "INTO" Then txtQuery = Replace(txtQuery, WordsArray(x), "INTO")
    If UCase(WordsArray(x)) = "AND" Then txtQuery = Replace(txtQuery, WordsArray(x), "AND")
    If UCase(WordsArray(x)) = "OR" Then txtQuery = Replace(txtQuery, WordsArray(x), "OR")
Next
txtQuery.SelStart = S
End Sub
Private Sub Winsock_Connect()
    Dim Data As String
    Data = "GET " & QueryFilePath & "?query=" & Replace(Replace(Replace(txtQuery, "'", "·"), " ", "%20"), vbCrLf, "") & " HTTP/1.1" & vbCrLf & _
            "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, application/vnd.ms-excel, application/msword, application/vnd.ms-powerpoint, */*" & vbCrLf & _
            "Accept-Encoding: deflate" & vbCrLf & _
            "User-Agent: ElFerPHPTunnel" & vbCrLf & _
            "Host: " & Host & vbCrLf & _
            "Connection: Keep-Alive" & vbCrLf & vbCrLf
    Winsock.SendData Data
End Sub
Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    Winsock.GetData Data
    txtResult = txtResult & Data
    Log = Data: Log.SelStart = Len(Log)
End Sub
