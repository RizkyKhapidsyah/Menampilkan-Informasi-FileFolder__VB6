VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menampilkan Informasi File/Folder"
   ClientHeight    =   2610
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  'Ganti direktory file di bawah ini
  'dengan nama folder/direktori atau file yang
  'Anda inginkan untuk ditampilkan informasinya...
  Call ShowFolderInfo(App.Path + "\Filenya-Disini")
End Sub

Sub ShowFolderInfo(foldername)
On Error GoTo Pesan
   Dim fs, f, s, k, l, m, n, o
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFolder(foldername) 'Direktori
   'Untuk info file, ganti GetFolder dengan GetFile...
   'Set f = fs.GetFile(foldername)  'File
   s = f.DateCreated  'Tanggal dibuat
   k = f.Size         'Ukuran isi folder/file
   l = f.Name         'Nama folder/file ybt
   m = f.Path         'Nama path lengkap (lokasi)
   n = f.Type         'Apakah folder atau file...
   'Tampilkan informasi folder/file...
   MsgBox "Tanggal & Jam dibuat: " & Format(s, "dd/mm/yyyy hh:mm:ss") & "" & _
          vbCrLf & "Ukuran = " & Format(k, "#,#") & " byte(s)" & vbCrLf & _
          "Nama folder/file = " & l & "" & vbCrLf & _
          "Path lengkap = " & m & "" & vbCrLf & _
          "Type = " & n & "", vbInformation, _
          "Informasi File/Folder: " & foldername
   Exit Sub
Pesan:   'Kemungkinan jika terjadi error...
  Select Case Err.Number
         Case 76
             MsgBox "Direktori atau file tidak ada!", _
                     vbCritical, "Error"
         Case Else
             MsgBox Err.Number & " - " _
                    & Err.Description
  End Select
End Sub

Sub ShowFileInfo(fileName)
On Error GoTo Pesan
   Dim fs, f, s, k, l, m, n, o
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile(fileName) 'Direktori
   'Untuk info file, ganti GetFolder dengan GetFile...
   'Set f = fs.GetFile(foldername)  'File
   s = f.DateCreated  'Tanggal dibuat
   k = f.Size         'Ukuran isi folder/file
   l = f.Name         'Nama folder/file ybt
   m = f.Path         'Nama path lengkap (lokasi)
   n = f.Type         'Apakah folder atau file...
   'Tampilkan informasi folder/file...
   MsgBox "Tanggal & Jam dibuat: " & Format(s, "dd/mm/yyyy hh:mm:ss") & "" & _
          vbCrLf & "Ukuran = " & Format(k, "#,#") & " byte(s)" & vbCrLf & _
          "Nama folder/file = " & l & "" & vbCrLf & _
          "Path lengkap = " & m & "" & vbCrLf & _
          "Type = " & n & "", vbInformation, _
          "Informasi File/Folder: " & fileName
   Exit Sub
Pesan:   'Kemungkinan jika terjadi error...
  Select Case Err.Number
         Case 76
             MsgBox "Direktori atau file tidak ada!", _
                     vbCritical, "Error"
         Case Else
             MsgBox Err.Number & " - " _
                    & Err.Description
  End Select
End Sub

Private Sub Command2_Click()
    Call ShowFileInfo(App.Path + "\Filenya-Disini\PilihSaya.txt")
End Sub

Private Sub Form_Load()
    Command1.Caption = "Tampilkan Informasi Folder"
    Command2.Caption = "Tampilkan Informasi File"
End Sub
