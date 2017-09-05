VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "密码"
   ClientHeight    =   4035
   ClientLeft      =   5040
   ClientTop       =   4620
   ClientWidth     =   6675
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   6675
   Begin VB.CommandButton Generate 
      Caption         =   "生成随机密钥"
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton showabout 
      Caption         =   "关于/帮助"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Encrypt 
      Caption         =   "加密"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Decrypt 
      Caption         =   "解密"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Check 
      Caption         =   "密钥验证"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Ciphertext 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox Plaintext 
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox Keytext 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "powered by ozem"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label PlainLabel 
      Caption         =   "明文"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label CipherLabel 
      Caption         =   "密文"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label keyLabel 
      Caption         =   "密钥"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Dim key(1 To 127) As String
Dim unkey(1 To 127) As String
Private Sub Check_Click()
Dim s As String, ch As String
Dim used(1 To 127) As Boolean
For i = 1 To 127
    used(i) = False
Next i
s = Keytext.Text
If Len(s) <> 26 Then
    MsgBox "密钥长度不合法"
    keyLabel.Caption = "密钥×"
    Exit Sub
End If
For i = 1 To 26
    ch = Mid(s, i, 1)
    If (used(Asc(ch))) Then
        MsgBox "密钥不合法"
        keyLabel.Caption = "密钥×"
        Exit Sub
    End If
    used(Asc(ch)) = True
    key(Asc("a") + i - 1) = ch
    unkey(Asc(ch)) = Chr(Asc("a") + i - 1)
Next i
flag = True
keyLabel.Caption = "密钥√"
End Sub

Private Sub Ciphertext_Change()
    CipherLabel.Caption = "密文..."
    PlainLabel.Caption = "明文"
End Sub

Private Sub Decrypt_Click()
Dim s As String, ch As String, t As String
If flag = False Then
    MsgBox "请输入并验证密钥"
    keyLabel.Caption = "密钥×"
    Exit Sub
End If
s = Ciphertext.Text
t = ""
If (Len(s) = 0) Then
    MsgBox "请输入密文"
    Exit Sub
End If
For i = 1 To Len(s)
    ch = Mid(s, i, 1)
    t = t + unkey(Asc(ch))
Next i
Plaintext.Text = t
PlainLabel.Caption = "明文√"
CipherLabel.Caption = "密文√"
End Sub

Private Sub Encrypt_Click()
Dim s As String, ch As String, t As String
s = Plaintext.Text
t = ""
If flag = False Then
    MsgBox "请输入并验证密钥"
    keyLabel.Caption = "密钥×"
    Exit Sub
End If
If Len(s) = 0 Then
    MsgBox "请输入明文"
    Exit Sub
End If
For i = 1 To Len(s)
    ch = Mid(s, i, 1)
    t = t + key(Asc(ch))
Next i
Ciphertext.Text = t
PlainLabel.Caption = "明文√"
CipherLabel.Caption = "密文√"
End Sub

Private Sub Form_Load()
For i = 1 To 127
    key(i) = Chr(i)
    unkey(i) = Chr(i)
Next i
End Sub

Private Sub Generate_Click()
'If flag Then Exit Sub
Randomize
Dim ch As String, s As String
Dim used(1 To 127) As Boolean
s = ""
For i = 1 To 127
    used(i) = False
Next i
For i = 1 To 26
    ch = Chr(Int(Rnd * 26) + Asc("a"))
    If (used(Asc(ch))) Then
        i = i - 1
    Else
    used(Asc(ch)) = True
    key(Asc("a") + i - 1) = ch
    unkey(Asc(ch)) = Chr(Asc("a") + i - 1)
    s = s + ch
    End If
Next i
Keytext.Text = s
keyLabel.Caption = "密钥√"
flag = True
End Sub

Private Sub Keytext_Change()
keyLabel.Caption = "密钥..."
flag = False
End Sub

Private Sub Plaintext_Change()
PlainLabel.Caption = "明文..."
CipherLabel.Caption = "密文"
End Sub

Private Sub showabout_Click()
frmAbout.Show
End Sub
