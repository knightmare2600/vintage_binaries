VERSION 4.00
Begin VB.Form frmSubettingCalc 
   Caption         =   "SubnetCalc"
   ClientHeight    =   3240
   ClientLeft      =   1245
   ClientTop       =   1710
   ClientWidth     =   4860
   ClipControls    =   0   'False
   Height          =   3645
   Icon            =   "FRMSUBET.frx":0000
   Left            =   1185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4860
   Top             =   1365
   Width           =   4980
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   2640
      TabIndex        =   16
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdOpenfile 
      Caption         =   "&Open Filre"
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   495
      Left            =   3600
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   12
      Text            =   "subnet.txt"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtHostQtyAnswer 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSubnetQty 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox TxtSubnetMaskAnswer 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox TxtOctet4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox TxtOctet3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox TxtOctet2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox TxtOctet1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label LbCopyright 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VB4 Subnet Calc by Knightmare 2024"
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label lblHostQty 
      Caption         =   "Filename:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblHostQty 
      Caption         =   "No. of devices:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblSubnetMaskAnswer 
      Caption         =   "No. of subnets:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label LblSubnetMask 
      Caption         =   "Subnet Mask:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblNetworkAddress 
      Caption         =   "Network Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmSubettingCalc"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
'Find the folder to write the subnet file to
'Dim FolderBrowser As New FolderBrowserDialog
FolderBrowser.ShowNewFolderButton = False
FolderBrowser.RootFolder = System.Environment.SpecialFolder.Desktop
FolderBrowser.SelectedPath = My.Computer.FileSystem.SpecialDirectories.Desktop
If FolderBrowser.ShowDialog = Windows.Forms.DialogResult.OK Then
Path = FolderBrowser.SelectedPath & "\"
End If
End Sub

Private Sub cmdCalculate_Click()
'Private Sub cmdCalculate_Click(ByVal sender As System.object, ByVal e As System.EventArgs)

'Public path As String
'Public newfile As String

'Handles cmdCalculate.Click

' declare variables
Dim octet1 As Double
Dim octet2 As Double
Dim octet3 As Double
Dim octet4 As Double
Dim subnet As Double

Dim subnet_last_octet As String
Dim netIP As String
Dim IpRange As String
Dim Broadcast As String
Dim counter As Double

'set variables
octet1 = TxtOctet1.Text
octet2 = txtOctet2.Text
octet3 = txtOctet3.Text
octet4 = txtOctet4.Text
subnet = TxtSubnetQty.Text
newfile = txtFilename.Text

' Urgh!
If subnet = 1 Then
  subnet_last_octet = "0"
  subnet = 1
End If
If subnet = 1 And subnet <= 2 Then
  subnet_last_octet = "128"
  subnet = 2
End If
If subnet > 2 And subnet <= 4 Then
  subnet_last_octet = "192"
  subnet = 4
End If
If subnet > 4 And subnet <= 8 Then
  subnet_last_octet = "224"
  subnet = 8
End If
If subnet > 8 And subnet <= 16 Then
  subnet_last_octet = "240"
  subnet = 16
End If
If subnet > 16 And subnet <= 32 Then
  subnet_last_octet = "248"
  subnet = 32
End If
If subnet > 32 And subnet <= 64 Then
  subnet_last_octet = "252"
  subnet = 64
End If
If subnet > 64 And subnet <= 128 Then
  subnet_last_octet = "254"
  subnet = 128
End If
If subnet > 128 And subnet <= 256 Then
  subnet_last_octet = "255"
  subnet = 256
End If

'set subnet mask for class A network
If octet1 >= 1 And octet1 <= 126 Then
  txtSubnetMaskAnswer.Text = "255." & subnet_last_octet & ".0.0"
  txtHostQtyAnswer.Text = 2 ^ 24 / subnet - 2
''  My.Computer.FileSystem.WriteAllText(path & newfile & ".txt"  "Subnet" & vbTab & "Network IP" & vbTab & vbTab &"IP Range" & vbTab & vbTab & vbTab & "Broadcast IP" & vbCrLf, True)
  counter = 0
  Do While counter < subnet
    netIP = Str(octet1) & "." & Str(octet2 + (256 / subnet) * counter) & "." & Str(octet3) & "." & Str(octet4)
    Broadcast = Str(octet1) & "." & Str(octet2 + (256 / subnet) * (counter + 1) - 1) & ". 255" & ". 255"
    IpRange = Str(octet1) & "." & Str(octet2 + (256 / subnet) * counter) & "." & Str(octet3) & "." & Str(octet4 + 1) & " - " & Str(octet1) & "." & Str(octet2 + (256 / subnet) * (counter + 1) - 1) & ". 255" & ". 254"
''    My.Computer.FileSystem.WriteAllText(path & newfile & ".txt", " " & Str(counter + 1) & vbTab & netIP & vbTab & IpRange & vbTab & Broadcast & vbCrLf, True)
    Open txtFilename.Text For Append As #1
    Write #1, Str(counter + 1) & vbTab & netIP & vbTab & IpRange & vbTab & Broadcast & vbCrLf
    counter = counter + 1
  Loop
End If

' set subnet mask for class B network
If octet1 >= 128 And octet1 <= 191 Then
  txtSubnetMaskAnswer.Text = "255.255." & subnet_last_octet & ".0"
  txtHostQtyAnswer.Text = 2 ^ 16 / subnet - 2
  counter = 0
  '' My.Computer.FileSystem.WriteAllText(path & newfile & ".txt", "Subnet" & vbTab & "Network IP"& vbTab & vbTab & "IP Range" & vbTab & vbTab & vbTab & "Broadcast IP" & vbCrLf, True)
  Do While counter < subnet
    netIP = Str(octet1) & "." & Str(octet2) & "." & Str(octet3 + (256 / subnet) * counter) & "." & Str(octet4)
    Broadcast = Str(octet1) & "." & Str(octet2) & "." & Str(octet3 + (256 / subnet) * (counter + 1) - 1) & ". 255"
    IpRange = Str(octet1) & "." & Str(octet2) & "." & Str(octet3 + (256 / subnet) * counter) & "." & Str(octet4 + 1) & " - " & Str(octet1) & "." & Str(octet2) & "." & Str(octet3 + (256 / subnet) * (counter + 1) - 1) & ". 254"
    '' My.Computer.FileSystem.WriteAllText(path & newfile & ".txt", " " & Str(counter + 1) & vbTab & netIP & vbTab & IpRange & vbTab & Broadcast & vbCrLf, True)
    Open txtFilename.Text For Append As #1
    Write #1, Str(counter + 1) & vbTab & netIP & vbTab & IpRange & vbTab & Broadcast & vbCrLf
    counter = counter + 1
  Loop
End If

'set subnet mask for class C network
If octet1 >= 192 And octet1 <= 223 Then
  txtSubnetMaskAnswer.Text = "255.255.255." & subnet_last_octet
  txtHostQtyAnswer.Text = (2 ^ 8) / (subnet - 2)
  counter = 0
  '' My.Computer.FileSystem.WriteAllText(path & newfile & ".txt", "Subnet" & vbTab & "Network IP" & vbTab & vbTab & "IP Range" & vbTab & vbTab & vbTab & "Broadcast IP" & vbCrLf, True)
  Do While counter < subnet
    netIP = Str(octet1) & "." & Str(octet2) & "." & Str(octet3) & "." & Str(octet4 + (256 / subnet) * counter)
    Broadcast = Str(octet1) & "." & Str(octet2) & "." & Str(octet3) & "." & Str(octet4 + (256 / subnet) * (counter + 1) - 1)
    IpRange = Str(octet1) & "." & Str(octet2) & "." & Str(octet3) & "." & Str(octet4 + (256 / subnet) * counter + 1) & " - " & Str(octet1) & "." & Str(octet2) & "." & Str(octet3) & "." & Str(octet4 + (256 / subnet) * (counter + 1) - 2)
    '' My.Computer.FileSystem.WriteAllText(path & newfile & ".txt", " " & Str(counter + 1) & vbTab & netIP & vbTab & IpRange& vbTab & Broadcast & vbCrLf, True)
    Open txtFilename.Text For Append As #1
    Write #1, Str(counter + 1) & vbTab & netIP & vbTab & IpRange & vbTab & Broadcast & vbCrLf
    counter = counter + 1
    Close #1
  Loop
End If

End Sub

Private Sub CmdExit_Click()
End
End Sub


Private Sub cmdOpenfile_Click()
'open the file to view and print
'System.Diagnostics.Process.Start("Notepad", path & newfile & ".txt")
End Sub

Private Sub cmdReset_Click()

' Clear textboxes, labels and path
TxtOctet1.Text = ""
txtOctet2.Text = ""
txtOctet3.Text = ""
txtOctet4.Text = ""
TxtSubnetQty.Text = ""
txtSubnetMaskAnswer.Text = ""
txtHostQtyAnswer.Text = ""
txtFilename.Text = ""
Path = ""

End Sub



Private Sub LbCopyright_Click()
' Easter Egg of sorts...
MsgBox "Kids today... Back in my day it was" & vbCrLf & "a flask of tea and a copy of vim!", vbExclamation, "Gaun yersel!"
End Sub



