VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Tratamento de dados"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Case.frx":0000
      Height          =   4695
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8281
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   8280
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=PostgreSQL30"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PostgreSQL30"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT* FROM pessoa"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Importar Dados"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exmportar Dados"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
'Exmportar dados

Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim strConn As String
strConn = "Driver={PostgreSQL Unicode};Server=localhost;Port=5432;Database=b.d_case;Uid=postgres;Pwd=1234;"

conn.Open strConn

rs.Open "SELECT * FROM pessoa", conn

Dim xlApp As New Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim i As Long, j As Long
For i = 0 To rs.Fields.Count - 1
   xlSheet.Cells(1, i + 1).Value = rs.Fields(i).Name
Next

i = 2
Do While Not rs.EOF
    For j = 0 To rs.Fields.Count - 1
        xlSheet.Cells(i, j + 1).Value = rs.Fields(j).Value
    Next j
    rs.MoveNext
    i = i + 1
Loop

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

Dim filePath As Variant
filePath = Application.GetSaveAsFilename(FileFilter:="Excel Workbook (*.xlsx), *.xlsx")

If filePath <> False Then
    xlBook.SaveAs filePath
    xlApp.Visible = False
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    MsgBox ("Arquivo salvo")
Else
    xlApp.Visible = False
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    MsgBox ("Arquivo não salvo")
End If

End Sub

Private Sub Command3_Click()
    ' Importar dados
    
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    Dim strConn As String
    strConn = "Driver={PostgreSQL Unicode};Server=localhost;Port=5432;Database=b.d_case;Uid=postgres;Pwd=1234;"
    
    conn.Open strConn
    
    Dim xl As Excel.Application
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    Dim i As Integer
    Dim sql As String
    
    Set xl = CreateObject("Excel.Application")
    
    Dim fileDialog As Object
    Set fileDialog = xl.fileDialog(3)
    fileDialog.AllowMultiSelect = False
    fileDialog.Title = "Selecione o arquivo XLSX"
    
   
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "Arquivos XLSX", "*.xlsx"
    
    
    If fileDialog.Show = -1 Then
        
        Dim filePath As String
        filePath = fileDialog.SelectedItems(1)
        
        Set wb = xl.Workbooks.Open(filePath)
        
        Set ws = wb.Worksheets(1)
        
        For i = 2 To ws.UsedRange.Rows.Count
            sql = "INSERT INTO pessoa (contrato, saldo, nome, endereço, telefone) VALUES (" & _
                "'" & ws.Cells(i, 1).Value & "', " & _
                "'" & ws.Cells(i, 2).Value & "', " & _
                "'" & ws.Cells(i, 3).Value & "', " & _
                "'" & ws.Cells(i, 4).Value & "', " & _
                "'" & ws.Cells(i, 5).Value & "')"
            conn.Execute sql
        Next i
        
        ' Reiniciar formulário
        Unload Me
        Load Form1
        Form1.WindowState = frmState
        Form1.Show
        
        conn.Close
        
        MsgBox "Arquivo importado!"
    Else
        MsgBox "Nenhum arquivo selecionado!"
    End If
End Sub


Private Sub btnVoltar_Click()
    Unload Me
    Form1.Show
End Sub


