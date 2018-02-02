VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form4"
   ScaleHeight     =   6765
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtkode 
      Height          =   405
      Left            =   3240
      TabIndex        =   10
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtnama 
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox cmbsatuan 
      Height          =   315
      ItemData        =   "Form4.frx":0000
      Left            =   3240
      List            =   "Form4.frx":0010
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtharga 
      Height          =   315
      Left            =   3240
      TabIndex        =   7
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtstock 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdubah 
      Caption         =   "ubah"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdcetak 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid grdtabel 
      Bindings        =   "Form4.frx":002C
      Height          =   1455
      Left            =   840
      TabIndex        =   0
      Top             =   3240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Daftar barang"
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
   Begin VB.Label Label1 
      Caption         =   "Pengelolaan Data Barang (sql)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Kode"
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Barang"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Satuan"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Harga"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Stock"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strkoneksi As String
Dim koneksidb As ADODB.Connection
Dim rsbarang As ADODB.Recordset
Dim strsql As String

Private Sub cmdbatal_Click()
txtkode = ""
txtnama = ""
cmbsatuan = ""
txtharga = ""
txtstock = ""
txtkode.SetFocus
End Sub

Private Sub cmdhapus_Click()
strsql = "delete from barang where kode ='" & txtkode & "'"
koneksidb.Execute strsql
'koneksidb.Execute strsql (untuk mengeksekusi)
rsbarang.Requery
txtkode.Locked = False
Call cmdbatal_Click
End Sub

Private Sub cmdsimpan_Click()
strsql = "INSERT INTO BARANG (kode,nama_barang,satuan,harga,stock) values " & _
"('" & txtkode & "','" & txtnama & "','" & cmbsatuan & "','" & txtharga & "','" & txtstock & "')"

koneksidb.Execute strsql
rsbarang.Requery

End Sub

Private Sub cmdubah_Click()
strsql = "update barang set nama_barang='" & txtnama & "',satuan='" & cmbsatuan & "',harga='" & txtharga & "',stock='" & txtstock & "' where kode='" & txtkode & "'"
koneksidb.Execute strsql

rsbarang.Requery
'rsbarang.Requery (untuk merefresh datagridnya)
txtkode.Locked = False


End Sub

Private Sub Form_Load()
strkoneksi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\AMIK\pemrograman visual basic\database\inventory.accdb;Persist Security Info=False"

'membuat koneksi ke database
Set koneksidb = New ADODB.Connection
koneksidb.Open strkoneksi
'(database sudah terhubung)

'membuat recordset barang
Set rsbarang = New ADODB.Recordset
rsbarang.CursorLocation = adUseClient
strsql = "select*from barang order by kode"
'(order by kode artinya data akan di urutkan secara otomatis dengan kode)
rsbarang.Open strsql, koneksidb

'tampilkan data barang pada datagrid
Set grdtabel.DataSource = rsbarang




End Sub

Private Sub grdtabel_DblClick()
txtkode = rsbarang!kode
txtnama = rsbarang!nama_barang
cmbsatuan = rsbarang!satuan
txtharga = rsbarang!harga
txtstock = rsbarang!stock
txtkode.Locked = True
End Sub
