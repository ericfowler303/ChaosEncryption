VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frm_encrypt 
   Caption         =   "Chaos Theory"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txt_output 
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmd_Decrypt 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txt_input 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txt_startingdata 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmd_encrypt 
      Caption         =   "Encrypt"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txt_key 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Output File"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Input File"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "0 to 1"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "0 to 4"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Starting data"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Encryption Key"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frm_encrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim key As Double

Private Sub cmd_Decrypt_Click()
'setting values to variables
key = Val(txt_key.Text)

'read file into byte array
Dim bytBuffer() As Byte
Dim strFileName As String
Dim lngFileNum As Long
Dim length_of_file As Double

'Open the file
numFile = FreeFile
strFileName = "mynewfile.ct"
Open strFileName For Binary As #numFile

'getting the starting data set
length_of_file = Val(LOF(1) - 1)
Dim chaosdata() As Double

'setting up the array of data to use
ReDim chaosdata(0 To length_of_file) As Double
chaosdata(0) = Val(txt_startingdata.Text)

For n = 0 To (length_of_file - 1)
    chaosdata(n + 1) = key * (1 - chaosdata(n)) * (chaosdata(n))
Next

'size the variable to be the same size as the file
    ReDim bytBuffer(LOF(1) - 1) As Byte
'load the data
    Get #numFile, , bytBuffer()

'close the file
Close #numFile

Dim fileNewData() As Integer
Dim fileWriteData() As Byte
ReDim fileNewData(length_of_file)
ReDim fileWriteData(length_of_file)

For i = 0 To length_of_file
    fileNewData(i) = (bytBuffer(i) - (chaosdata(i) * 255))

    If fileNewData(i) >= 0 Then
        fileWriteData(i) = fileNewData(i)
        Else
        fileWriteData(i) = (fileNewData(i) + 256)
    End If
ProgressBar1.Value = (i / length_of_file) * 100
Next

'Create new encoded file using new data
    numFile = FreeFile
    encFileName = "decrypted-file.wmv"
    Open encFileName For Binary As #numFile

    'Write the contents of the variable to the file
    Put #numFile, , fileWriteData()

    'close the file
    Close #numFile
End Sub

Private Sub cmd_encrypt_Click()
'setting values to variables
key = Val(txt_key.Text)

'read file into byte array
Dim bytBuffer() As Byte
Dim strFileName As String
Dim lngFileNum As Long
Dim length_of_file As Double

'Open the file
numFile = FreeFile
strFileName = "jetpack.wmv"
Open strFileName For Binary As #numFile

'getting the starting data set
length_of_file = Val(LOF(1) - 1)
Dim chaosdata() As Double
ReDim chaosdata(0 To length_of_file) As Double
chaosdata(0) = Val(txt_startingdata.Text)

'setting up the array of data to use
For n = 0 To (length_of_file - 1)
    chaosdata(n + 1) = key * (1 - chaosdata(n)) * (chaosdata(n))
Next

'size the variable to be the same size as the file
    ReDim bytBuffer(LOF(1) - 1) As Byte
'load the data
    Get #numFile, , bytBuffer()

'close the file
Close #numFile

Dim fileEncrypt() As Byte
Dim fileWriteData() As Byte
Dim fileNewData() As Integer

ReDim fileEncrypt(length_of_file)
ReDim fileNewData(length_of_file)
ReDim fileWriteData(length_of_file)

For i = 0 To length_of_file
    fileEncrypt(i) = ((chaosdata(i) * 255) + bytBuffer(i)) Mod 256
    fileNewData(i) = (fileEncrypt(i) - (chaosdata(i) * 255))
    If fileNewData(i) >= 0 Then
        fileWriteData(i) = fileNewData(i)
        Else
        fileWriteData(i) = (fileNewData(i) + 256)
    End If
    
    If fileWriteData(i) <> bytBuffer(i) Then MsgBox "Error in encryption!"
        
    ProgressBar1.Value = (i / length_of_file) * 100
Next
'Create new encoded file using new data
    numFile = FreeFile
    encFileName = "MyNewFile.ct"
    Open encFileName For Binary As #numFile

    'Write the contents of the variable to the file
    Put #numFile, , fileEncrypt()

    'close the file
    Close #numFile

End Sub

Private Sub Form_Load()
'centering the form
frm_encrypt.Left = (Screen.Width / 2) - (frm_encrypt.Width / 2)
frm_encrypt.Top = (Screen.Height / 2) - (frm_encrypt.Height / 2)
ProgressBar1.Value = 0

End Sub
