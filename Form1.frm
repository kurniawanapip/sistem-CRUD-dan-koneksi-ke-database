VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adobarang 
      Height          =   375
      Left            =   4440
      Top             =   5400
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":0087
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "barang"
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
   Begin MSDataGridLib.DataGrid grdtabel 
      Bindings        =   "Form1.frx":010E
      Height          =   1455
      Left            =   480
      TabIndex        =   16
      Top             =   3600
      Width           =   6015
      _ExtentX        =   10610
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
   Begin VB.CommandButton cmdcetak 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdubah 
      Caption         =   "ubah"
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtstock 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtharga 
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.ComboBox cmbsatuan 
      Height          =   315
      ItemData        =   "Form1.frx":0126
      Left            =   2280
      List            =   "Form1.frx":0136
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtnama 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtkode 
      Height          =   405
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Stock"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Harga"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Satuan"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Barang"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Kode"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Pengelolaan Data Barang (Ado)"
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
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vhasilcek As Boolean

Private Sub cmdbatal_Click()
txtkode.Locked = False
txtkode.Text = ""
txtnama.Text = ""
cmbsatuan.Text = ""
txtharga.Text = ""
txtstock.Text = ""
txtkode.SetFocus
End Sub

Private Sub cmdcetak_Click()
Form2.Show
End Sub

Private Sub cmdhapus_Click()
Adobarang.Recordset.MoveFirst
Adobarang.Recordset.Find "Kode='" & txtkode.Text & " ' "
If Not Adobarang.Recordset.EOF Then
 Adobarang.Recordset.Delete
 Adobarang.Recordset.Update
End If
End Sub

Private Sub cmdsimpan_Click()
Call cekkode
If vhasilcek Then
 Adobarang.Recordset.AddNew
 Adobarang.Recordset.Fields("Kode") = txtkode.Text
 Adobarang.Recordset.Fields("Nama_Barang") = txtnama.Text
 Adobarang.Recordset.Fields("Satuan") = cmbsatuan.Text
 Adobarang.Recordset.Fields("Harga") = txtharga.Text
 Adobarang.Recordset.Fields("Stock") = txtstock.Text
 Adobarang.Recordset.Update
End If
End Sub

Private Sub cmdubah_Click()
Adobarang.Recordset.MoveFirst
Adobarang.Recordset.Find "Kode='" & txtkode.Text & " ' "
If Not Adobarang.Recordset.EOF Then
 Adobarang.Recordset.Fields("Kode") = txtkode.Text
 Adobarang.Recordset.Fields("Nama_Barang") = txtnama.Text
 Adobarang.Recordset.Fields("Satuan") = cmbsatuan.Text
 Adobarang.Recordset.Fields("Harga") = txtharga.Text
 Adobarang.Recordset.Fields("Stock") = txtstock.Text
 Adobarang.Recordset.Update
End If

End Sub

Private Sub grdtabel_DblClick()
txtkode.Text = Adobarang.Recordset.Fields("kode")
txtnama.Text = Adobarang.Recordset.Fields("Nama_Barang")
cmbsatuan.Text = Adobarang.Recordset.Fields("Satuan")
txtharga.Text = Adobarang.Recordset.Fields("Harga")
txtstock.Text = Adobarang.Recordset.Fields("Stock")
txtkode.Locked = True
End Sub

Sub cekkode()
If Not Adobarang.Recordset.BOF Then
 Adobarang.Recordset.MoveFirst
 Adobarang.Recordset.Find "Kode='" & txtkode.Text & " ' "
 If Adobarang.Recordset.EOF Then
   vhasilcek = True
 Else
   MsgBox "kode barang sudah ada."
   vhasilcek = False
 End If
Else
 vhasilcek = True
End If

End Sub
