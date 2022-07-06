VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Registry Example"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Startup Folder"
      Height          =   390
      Left            =   4425
      TabIndex        =   5
      Top             =   630
      Width           =   2085
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Long"
      Height          =   390
      Left            =   2265
      TabIndex        =   4
      Top             =   630
      Width           =   2085
   End
   Begin VB.CommandButton Command3 
      Caption         =   "String"
      Height          =   390
      Left            =   90
      TabIndex        =   3
      Top             =   1080
      Width           =   2085
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Key"
      Height          =   390
      Left            =   2265
      TabIndex        =   2
      Top             =   1080
      Width           =   2085
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   210
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Byte Array"
      Height          =   390
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Reg As Registry
Attribute Reg.VB_VarHelpID = -1



Private Sub Command1_Click()
 Dim mydata(20) As Byte
 Dim BlankArray(0) As Byte
 Dim x As Byte

  Reg.hKey = HKEY_LOCAL_MACHINE 'set Key
  Reg.DataType = REG_BINARY 'set data type
  Reg.SubKey = "tester" 'set sub key
  Reg.ValueName = "ByteArray" 'set value name

 For x = 1 To 20 'populate an array
  mydata(x) = x
 Next x

 Reg.Data = mydata 'set data

 Reg.SaveSetting 'save data to registry

 Reg.Data = BlankArray 'set data to blank array to verify read

 Reg.GetSetting 'get data

 For x = 0 To UBound(Reg.Data) 'put data in the text box
  Text1.Text = Text1.Text & ":" & Reg.Data(x)
 Next x

End Sub

Private Sub Command2_Click()
 Reg.hKey = HKEY_LOCAL_MACHINE 'deletes a key
 Reg.SubKey = "tester"
 Reg.DeleteKey
End Sub

Private Sub Command3_Click()
 Reg.hKey = HKEY_LOCAL_MACHINE
 Reg.DataType = REG_SZ
 Reg.SubKey = "tester"
 Reg.ValueName = "String"

 Reg.Data = "Karl Grear"

 Reg.SaveSetting

 Reg.Data = "hi there"

 Reg.GetSetting

 Text1.Text = Reg.Data
End Sub

Private Sub Command4_Click()

 Reg.hKey = HKEY_LOCAL_MACHINE
 Reg.DataType = REG_DWORD
 Reg.SubKey = "tester"
 Reg.ValueName = "Long"

 Reg.Data = 123456

 Reg.SaveSetting

 Reg.Data = 55555

 Reg.GetSetting

 Text1.Text = Reg.Data
End Sub

Private Sub Command5_Click()
 Reg.hKey = HKEY_CURRENT_USER
 Reg.SubKey = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" 'example of rooted sub keys
 Reg.ValueName = "Startup"
 Reg.DataType = REG_SZ
 Reg.GetSetting
 Text1.Text = Reg.Data
End Sub

Private Sub Form_Load()
 Set Reg = New Registry

End Sub

Private Sub Reg_onFailed()
 MsgBox "aawhh shoot "
End Sub

Private Sub Reg_onSuccess()
 MsgBox "It worked!! :)"
End Sub

