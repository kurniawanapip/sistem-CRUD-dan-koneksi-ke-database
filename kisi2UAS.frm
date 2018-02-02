VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00004040&
   Caption         =   "hitung gaji"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form5"
   ScaleHeight     =   5745
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Timer Timer3 
      Interval        =   30
      Left            =   120
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   1920
   End
   Begin MSDataGridLib.DataGrid grdtabel 
      Height          =   1575
      Left            =   1200
      TabIndex        =   8
      Top             =   3480
      Width           =   9375
      _ExtentX        =   16536
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
   Begin VB.CommandButton cmdtampil 
      Caption         =   "&Tampilkan Gaji Karyawan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan Data Gaji Karyawan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   1680
      Width           =   3255
   End
   Begin VB.ComboBox cmbgol 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      ItemData        =   "kisi2UAS.frx":0000
      Left            =   3240
      List            =   "kisi2UAS.frx":000D
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtnama 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   1320
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004040&
      Caption         =   "happy coding.!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   9
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Golongan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Nama Karyawan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   8655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11055
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strkoneksi As String
Dim koneksidb As ADODB.Connection
Dim rskaryawan As ADODB.Recordset
Dim rsgajipokok As ADODB.Recordset
Dim rsgrid As ADODB.Recordset
Dim sql As String
Dim i As Long
Dim merah, hijau, biru As Integer
'Dim vgapok, vtun, vpajak, vgaji, vins As Single

Private Sub cmdsimpan_Click()
 Set rskaryawan = New ADODB.Recordset
sql = "insert into karyawan (Nama_Karyawan,Gol) VALUES " & _
      "('" & txtnama.Text & "','" & cmbgol.Text & "')"
koneksidb.Execute sql

MsgBox "Data Berhasil Di simpan", vbOKOnly + vbInformation, "informasi"
txtnama = ""
cmbgol = ""
txtnama.SetFocus
End Sub

Private Sub cmdtampil_Click()
ProgressBar1.Visible = True
Timer3.Enabled = True
End Sub

Private Sub Command1_Click()
sql = "delete from karyawan where Nama_Karyawan ='" & txtnama & "'"
koneksidb.Execute sql

rsgrid.Delete
txtnama = ""
cmbgol = ""
txtnama.SetFocus
End Sub

Private Sub Form_Load()
i = 0
Label2.FontSize = 20
 Timer3.Enabled = False
Label2 = "PENGELOLAAN GAJI KARYAWAN PT.SUDO"

strkoneksi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\AMIK\pemrograman visual basic\database\inventory.accdb;Persist Security Info=False"

Set koneksidb = New ADODB.Connection
koneksidb.Open strkoneksi

Call tabelgrid
End Sub

Private Sub grdtabel_Click()
txtnama.Text = rsgrid!Nama_Karyawan
cmbgol.Text = rsgrid!Gol

End Sub

Private Sub Timer1_Timer()
Label2.ForeColor = RGB(Rnd * 250, Rnd * 250, Rnd * 250)
    If (Label2.Left + Label2.Width) <= 0 Then
        Label2.Left = Me.Width
    End If
    Label2.Left = Label2.Left - 100
End Sub

Private Sub Timer2_Timer()
i = i + 1
If i = 1000000 Then i = 0 'Supaya tdk overflow, dsb...
merah = Int(255 * Rnd) 'Bangkitkan angka random untuk merah
hijau = Int(255 * Rnd) 'Bangkitkan angka random untuk hijau
biru = Int(255 * Rnd) 'Bangkitkan angka random untuk biru
Label5.ForeColor = RGB(merah, hijau, biru) 'Campur tiga warna
If i Mod 2 = 0 Then 'Jika counter habis dibagi 2
Label5.Visible = True 'Tampilkan label
Label5.Visible = False 'Sembunyikan label
Else 'Jika counter tidak habis dibagi 2
End If 'Akhir pemeriksaan
End Sub
Sub hitung()
 If rsgrid!Gol = "C" Then
 rsgrid!Insentif = "500000"
 Else
 If rsgrid!Gol = "B" Then
 rsgrid!Insentif = "200000"
 Else
 If rsgrid!Gol = "A" Then
 rsgrid!Insentif = "200000"
 End If
 End If
 End If
End Sub

Sub tabelgrid()
Set rsgrid = New ADODB.Recordset
rsgrid.Fields.Append "Nama_Karyawan", adVarChar, 30
rsgrid.Fields.Append "Gol", adVarChar, 1
rsgrid.Fields.Append "Gapok", adSingle
rsgrid.Fields.Append "Tunjangan", adSingle
rsgrid.Fields.Append "Insentif", adSingle
rsgrid.Fields.Append "PPH", adSingle
rsgrid.Fields.Append "Gaji", adSingle
rsgrid.Open

For X = 1 To 30
   rsgrid.AddNew
   Next
   rsgrid.MoveFirst
   
   Set grdtabel.DataSource = rsgrid
   

End Sub

Private Sub Timer3_Timer()
 ProgressBar1.Value = ProgressBar1.Value + 1
     If ProgressBar1.Value = ProgressBar1.Max Then
       Call proses
       MsgBox "complete", vbInformation, "informasi"
       Timer3.Enabled = False
       ProgressBar1.Visible = False
     End If
End Sub
Sub proses()
Set rskaryawan = New ADODB.Recordset
 sql = "select * from karyawan inner join gajipokok on karyawan.Gol = gajipokok.Gol"
rskaryawan.Open sql, koneksidb

Call tabelgrid
rskaryawan.MoveFirst
Do While Not rskaryawan.EOF
    rsgrid!Nama_Karyawan = rskaryawan!Nama_Karyawan
    rsgrid!Gol = rskaryawan.Fields("karyawan.Gol")
    rsgrid!Gapok = rskaryawan!Gapok
    Call hitung
    rsgrid!Tunjangan = rsgrid!Gapok * 20 / 100
    rsgrid!PPH = 15 / 100 * (rsgrid!Gapok + rsgrid!Tunjangan + rsgrid!Insentif)
    rsgrid!Gaji = (rsgrid!Gapok + rsgrid!Tunjangan + rsgrid!Insentif) - rsgrid!PPH
    rskaryawan.MoveNext
    rsgrid.MoveNext
    Loop
End Sub

