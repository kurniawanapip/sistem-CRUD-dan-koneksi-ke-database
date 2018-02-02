VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00004040&
   Caption         =   "Form6"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form6"
   ScaleHeight     =   4845
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdtampil 
      Caption         =   "tampil tagihan pelanggan"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   4080
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid grdtabel 
      Height          =   3255
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5741
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
      Caption         =   "DAFTAR TAGIHAN LISTRIK PELANGGAN"
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
            LCID            =   1033
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
            LCID            =   1033
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
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strkoneksi As String
Dim koneksidb As ADODB.Connection
Dim rspelanggan As ADODB.Recordset
Dim rstdl As ADODB.Recordset
Dim rsgrid As ADODB.Recordset
Dim sql As String
Dim vtarif, vjamnyala As Single
Private Sub cmdtampil_Click()
Set rspelanggan = New ADODB.Recordset
 sql = "select * from pelanggan inner join tdl on pelanggan.Daya = tdl.Daya"
rspelanggan.Open sql, koneksidb

Call tabelgrid
rspelanggan.MoveFirst
Do While Not rspelanggan.EOF
    rsgrid!ID = rspelanggan!ID
    rsgrid!Nama = rspelanggan!Nama
    rsgrid!Daya = rspelanggan.Fields("pelanggan.Daya")
    rsgrid!Kwh = rspelanggan!Kwh
    Call hitung
    rsgrid!Biaya_Beban = rspelanggan!Beban * (rsgrid!Daya / 1000)
    rsgrid!Biaya_Pemakaian = vtarif * rsgrid!Kwh
    rsgrid!PPJ = 5 / 100 * (rsgrid!Biaya_Beban + rsgrid!Biaya_Pemakaian)
    rsgrid!Tagihan = rsgrid!Biaya_Beban + rsgrid!Biaya_Pemakaian + rsgrid!PPJ
    rspelanggan.MoveNext
    rsgrid.MoveNext
    Loop

End Sub

Private Sub Form_Load()
strkoneksi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\AMIK\pemrograman visual basic\database\inventory.accdb;Persist Security Info=False"

Set koneksidb = New ADODB.Connection
koneksidb.Open strkoneksi

Call tabelgrid

End Sub
Sub tabelgrid()
Set rsgrid = New ADODB.Recordset
rsgrid.Fields.Append "ID", adVarChar, 2
rsgrid.Fields.Append "Nama", adVarChar, 30
rsgrid.Fields.Append "Daya", adSingle
rsgrid.Fields.Append "Kwh", adSingle
rsgrid.Fields.Append "Biaya_Beban", adSingle
'rsgrid.Fields.Append "jam", adSingle
rsgrid.Fields.Append "Biaya_Pemakaian", adSingle
rsgrid.Fields.Append "PPJ", adSingle
rsgrid.Fields.Append "Tagihan", adSingle
rsgrid.Open

For X = 1 To 30
   rsgrid.AddNew
   Next
   rsgrid.MoveFirst
   
   Set grdtabel.DataSource = rsgrid
End Sub
Sub hitung()
vjamnyala = rsgrid!Kwh / (rsgrid!Daya / 1000)
 
  If rsgrid!Daya = 450 Then
   vtarif = 400
  Else
  If rsgrid!Daya = 900 Then
   vtarif = 600
  Else
  If rsgrid!Daya = 1300 Then
   If vjamnyala < 40 Then
    vtarif = 950 - 50
    Else
    vtarif = 950
 End If
 End If
 End If
End If
End Sub

