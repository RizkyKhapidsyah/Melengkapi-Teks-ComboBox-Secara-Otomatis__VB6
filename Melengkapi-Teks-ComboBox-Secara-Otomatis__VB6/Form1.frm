VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Melengkapi Teks di Combobox secara Otomatis"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Text            =   "Combo5"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Text            =   "Combo4"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Text            =   "Combo3"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Konstanta untuk membantu pencarian string
Const CB_FINDSTRING = &H14C

Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long

'Untuk membantu menentukan apakah terjadi perubahan '(Ubah)hasil string di combobox atau tidak (Asli).
Public Enum EnumKarakter
  Asli = 0
  Ubah = 1
End Enum
Dim Karakter As EnumKarakter

'Ini untuk mengisi setiap combobox dengan data yang 'sama.
'Perhatikan perbedaan hasilnya saat data diketikkan di
'masing2 combobox ybt pada event procedure KeyPress...
Private Sub IsiSemuaCombobox()
  Dim ctrl As Control
  For Each ctrl In Form1.Controls
    If TypeOf ctrl Is ComboBox Then
      With ctrl
      .AddItem "Rizky Khapidsyah"
      .AddItem "Zubair Ahmad"
      .AddItem "Murtala Rizky"
      .AddItem "Sugiono"
      .AddItem "Sarwoto"
      .AddItem "Suherman"
      .AddItem "Sariani"
      .Text = .List(0)
      End With
    End If
  Next
End Sub

'Mula-mula, isi semua combobox dengan data yang sama
Private Sub Form_Load()
  IsiSemuaCombobox
End Sub

'Bandingkan Combo1 dan Combo4...
'Hasilnya sama saja bukan? Karena huruf yang akan 'digunakan tidak terpengaruh kepada parameter ketiga '(bUpperCase), tapi ditentukan oleh parameter keempat '(cCharacter), yang bernilai "Asli", artinya 'menggunakan karakter aslinya.

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  KeyAscii = AutoComplete(Combo1, KeyAscii, True, Asli)
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
  KeyAscii = AutoComplete(Combo4, KeyAscii, False, _
             Asli)
End Sub

'Karena parameter ketiga = False dan parameter keempat 'di-Ubah, maka huruf yang ditampilkan saat diketik akan 'menjadi huruf kecil semuanya (terjadi perubahan karena "Ubah").

Private Sub Combo2_KeyPress(KeyAscii As Integer)
  KeyAscii = AutoComplete(Combo2, KeyAscii, False, _
             Ubah)
End Sub

'Karena parameter ketiga = True, dan parameter keempat 'di-Ubah, maka huruf yang ditampilkan saat diketik akan 'menjadi huruf besar semuanya (terjadi perubahan karena '"Ubah").

Private Sub Combo3_KeyPress(KeyAscii As Integer)
  KeyAscii = AutoComplete(Combo3, KeyAscii, True, Ubah)
End Sub

'Karena parameter ketiga dan keempat tidak 'didefinisikan secara eksplisit dalam pemakaiannya, 'maka akan menggunakan parameter default-nya; masing-'masing: True dan Asli, sehingga huruf yang ditampilkan 'menjadi apa adanya (Asli).
'Dalam hal ini, sama dengan Combo1 dan Combo4 di atas.
Private Sub Combo5_KeyPress(KeyAscii As Integer)
  KeyAscii = AutoComplete(Combo5, KeyAscii)
End Sub

Public Function AutoComplete( _
       cbCombo As ComboBox, _
       sKeyAscii As Integer, _
       Optional bUpperCase As Boolean = True, _
       Optional cCharacter As EnumKarakter = Asli) _
       As Integer
  Dim lngFind As Long, intPos As Integer
  Dim intLength As Integer, tStr As String
  With cbCombo
    If sKeyAscii = 8 Then
       If .SelStart = 0 Then Exit Function
       .SelStart = .SelStart - 1
       .SelLength = 32000
       .SelText = ""
    Else
       'simpan posisi kursor
       intPos = .SelStart
       'simpan data string
       tStr = .Text
       'If bUpperCase = Asli Then
       '   .SelText = (Chr(sKeyAscii))
       If bUpperCase = True Then
          'ganti string. (hanya huruf besar)
          .SelText = UCase(Chr(sKeyAscii))
       Else 'If bUpperCase = KecilSemua Then
          'ganti string. (biarkan data apa adanya)
          .SelText = (Chr(sKeyAscii))
       End If
    End If
    'Cari string di combobox
    lngFind = SendMessage(.hwnd, CB_FINDSTRING, 0, _
              ByVal .Text)
    If lngFind = -1 Then 'Jika string tidak ditemukan
      'Set ke string yg lama (digunakan untuk data yang
      'membutuhkan pengawasan karakter
       .Text = tStr
       'Tentukan posisi kursor
       .SelStart = intPos
       'Tentukan panjang yang terpilih
       .SelLength = (Len(.Text) - intPos)
       'Kembalikan nilai 0 KeyAscii (tidak melakukan
       'apapun)
       AutoComplete = 0
       Exit Function
    Else 'Jika string ditemukan
       intPos = .SelStart 'Simpan posisi kursor
       'Simpan panjang teks sisa yang disorot
       intLength = Len(.List(lngFind)) - Len(.Text)
       If cCharacter = Ubah Then
        'Ganti teks baru dalam string (ubah seluruhnya)
         .SelText = .SelText & Right(.List(lngFind), _
                    intLength)
       Else  'Asli, huruf asli yang digunakan, tidak
             'diganti
         .Text = .List(lngFind)
       End If
       'Tentukan posisi kursor
       .SelStart = intPos
       'Tentukan panjang yang terpilih
       .SelLength = intLength
    End If
  End With
End Function


Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
