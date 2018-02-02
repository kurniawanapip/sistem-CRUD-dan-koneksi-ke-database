VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00008080&
   Caption         =   "Form3"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form3"
   ScaleHeight     =   5160
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdhapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdubah 
      Caption         =   "Ubah"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   4320
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid grdtabel 
      Height          =   1575
      Left            =   480
      TabIndex        =   10
      Top             =   2040
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2778
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   4080
      TabIndex        =   7
      Top             =   840
      Width           =   2895
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Total"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.TextBox txtkasir 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dtptgl 
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   89718785
      CurrentDate     =   42853
   End
   Begin MSDataListLib.DataCombo cmbnota 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000010&
      Caption         =   "Kasir"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000010&
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000010&
      Caption         =   "No. Nota"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Relasi tabel database"
      BeginProperty Font 
         Name            =   "Cooper Std Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strkoneksi As String
Dim koneksidb As ADODB.Connection
Dim rsgrid As ADODB.Recordset
Dim rsbarang As ADODB.Recordset
Dim sql As String
Dim vtotal, vjumlah As Single
Dim rsnota As ADODB.Recordset
Dim rsjual As ADODB.Recordset

Private Sub cmbnota_Click(Area As Integer)
If Area = dbcAreaList Then
'ambil data dari tabel jual
  Set rsjual = New ADODB.Recordset
  sql = "select * from jual where no_nota='" & cmbnota.Text & "'"
  rsjual.Open sql, koneksidb
  
  vtotal = rsjual!total_jual
  lbltotal.Caption = Format(vtotal, "currency")
  dtptgl.Value = rsjual!tgl_nota
  txtkasir.Text = rsjual!kasir
  
'ambil datadetail dari tabel jualdetail
 Set rsjual = New ADODB.Recordset
 sql = "select * from barang inner join jualdetail on barang.kode = jualdetail.kode " & _
       "where no_nota = '" & cmbnota.Text & "'"
rsjual.Open sql, koneksidb

Call tabelgrid
rsjual.MoveFirst
Do While Not rsjual.EOF
    rsgrid!kode = rsjual.Fields("barang.kode")
    rsgrid!nama_barang = rsjual!nama_barang
    rsgrid!harga = rsjual!harga
    rsgrid!qty = rsjual!qty
    rsgrid!jumlah = rsjual!jumlah
       
    rsjual.MoveNext
    rsgrid.MoveNext
    Loop
    
End If
End Sub

Private Sub cmdbatal_Click()
cmbnota.Text = ""
txtkasir.Text = ""
Call tabelgrid
cmbnota.SetFocus
End Sub

Private Sub cmdhapus_Click()
sql = "delete from jual where no_nota='" & cmbnota.Text & "'"
koneksidb.Execute sql

 Set rsjual = New ADODB.Recordset
 sql = "delete * from jualdetail " & _
       "where no_nota = '" & cmbnota.Text & "'"
koneksidb.Execute sql

Call tabelgrid
'rsjual.MoveFirst
'Do While Not rsjual.EOF
 '   rsgrid!kode = rsjual.Fields("barang.kode")
  '  rsgrid!nama_barang = rsjual!nama_barang
   ' rsgrid!harga = rsjual!harga
    'rsgrid!qty = rsjual!qty
    'rsgrid!jumlah = rsjual!jumlah
       
    'rsjual.MoveNext
    'rsgrid.MoveNext
    'Loop
cmbnota.Text = ""
txtkasir.Text = ""
Call tabelgrid
cmbnota.SetFocus
MsgBox "Data Berhasil Di hapus", vbOKOnly + vbInformation, "informasi"
End Sub

Private Sub cmdsimpan_Click()
'simpan data ke tabel jual
sql = "insert into jual(no_nota,tgl_nota,kasir,total_jual) VALUES " & _
      "('" & cmbnota.Text & "','" & dtptgl.Value & "','" & txtkasir & "','" & vtotal & "')"
koneksidb.Execute sql

'simpan detail penjualan kedalam tabel jualdetail
rsgrid.MoveFirst
Do While rsgrid!kode <> ""
 sql = "insert into jualdetail (no_nota,kode,qty,jumlah) values " & _
     "('" & cmbnota.Text & "','" & rsgrid!kode & "','" & rsgrid!qty & "','" & rsgrid!jumlah & "')"
     koneksidb.Execute sql
     rsgrid.MoveNext
Loop
MsgBox "data telah di simpan"
rsnota.Requery
cmbnota.Text = ""
txtkasir.Text = ""
Call tabelgrid
cmbnota.SetFocus
End Sub

Private Sub cmdubah_Click()
'sql = "update rsnota"
'koneksidb.Execute sql

'simpan detail penjualan kedalam tabel jualdetail

'rsgrid.MoveFirst
'Do While rsgrid!kode <> ""
 'sql = "update jualdetail set (no_nota,kode,qty,jumlah) values " & _
     "('" & cmbnota.Text & "','" & rsgrid!kode & "','" & rsgrid!qty & "','" & rsgrid!jumlah & "')"
  '   koneksidb.Execute sql
   '  rsgrid.MoveNext
'rsgrid.Update
'Loop
'MsgBox "data telah di ubah"
'rsnota.Requery
End Sub

Private Sub Form_Load()
strkoneksi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\AMIK\pemrograman visual basic\database\inventory.accdb;Persist Security Info=False"

'membuat koneksi ke database
Set koneksidb = New ADODB.Connection
koneksidb.Open strkoneksi
'(database sudah terhubung)
Call tabelgrid

'mengisi data combo dengan tabel jual
Set rsnota = New ADODB.Recordset
sql = "SELECT no_nota from jual"
rsnota.Open sql, koneksidb, adOpenKeyset

Set cmbnota.RowSource = rsnota 'muncul di datacombo nota
cmbnota.ListField = "no_nota"

End Sub

Sub tabelgrid()
Set rsgrid = New ADODB.Recordset
rsgrid.Fields.Append "Kode", adVarChar, 3
rsgrid.Fields.Append "Nama_Barang", adVarChar, 30
rsgrid.Fields.Append "Harga", adSingle
rsgrid.Fields.Append "Qty", adSingle
rsgrid.Fields.Append "Jumlah", adSingle
rsgrid.Open

For X = 1 To 20
   rsgrid.AddNew
   Next
   rsgrid.MoveFirst
   
   Set grdtabel.DataSource = rsgrid
   

End Sub

Sub caribarang()
Set rsbarang = New ADODB.Recordset
sql = "SELECT * FROM barang WHERE kode='" & rsgrid!kode & "'"
rsbarang.Open sql, koneksidb

If Not rsbarang.EOF Then
 rsgrid!kode = rsbarang!kode
 rsgrid!nama_barang = rsbarang!nama_barang
 rsgrid!harga = rsbarang!harga
 rsgrid!qty = 1
 rsgrid!jumlah = rsgrid!harga * rsgrid!qty

'hitung total
 vtotal = vtotal + rsgrid!jumlah
 lbltotal.Caption = Format(vtotal, "currency")
 rsgrid.MoveNext
Else
 MsgBox "kode barang tidak di temukan."
 rsgrid!kode = ""
End If

End Sub

Private Sub grdtabel_AfterColEdit(ByVal ColIndex As Integer)
vjumlah = rsgrid!jumlah
Select Case ColIndex
Case 0:
   Call caribarang
Case 3:
 Call rubahqty
End Select

End Sub

Sub rubahqty()
rsgrid!jumlah = rsgrid!harga * rsgrid!qty
vtotal = vtotal + (rsgrid!harga * rsgrid!qty) - vjumlah
lbltotal.Caption = Format(vtotal, "currency")
End Sub
