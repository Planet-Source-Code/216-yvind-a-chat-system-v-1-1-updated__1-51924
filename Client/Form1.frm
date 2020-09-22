VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Disconnect"
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtMotta 
      Height          =   2595
      Left            =   60
      TabIndex        =   8
      Top             =   420
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4577
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.ListBox lstNicks 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   3180
      TabIndex        =   7
      Top             =   420
      Width           =   1155
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0082
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.TextBox txtNick 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Text            =   "NoName"
         Top             =   60
         Width           =   1275
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   420
         TabIndex        =   4
         Text            =   "127.0.0.1"
         Top             =   60
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Nickname:"
         Height          =   195
         Left            =   2000
         TabIndex        =   6
         Top             =   700
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   3180
      TabIndex        =   2
      Top             =   3060
      Width           =   615
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   3060
      Width           =   3075
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   3840
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   344
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "23:16"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&


Private Sub cmdSend_Click()
If Not txtSend.Text = "" Then ' If the txtbox is not empty
    WS.SendData " <" & txtNick.Text & ">" & " " & txtSend.Text + vbCrLf ' Sends the nick and the text in the send-box
    txtSend.Text = "" ' Makes the txtbox empty and ready for new studd to send..
End If
End Sub

Private Sub Command1_Click()
WS.SendData "REM" & " " & txtNick.Text + vbCrLf ' Sends "REM" if you disconnect
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If WS.State = 7 Then ' If connected
    WS.SendData "REM" & " " & txtNick.Text + vbCrLf ' Sends "REM" if you disconnect
    'WS.SendData "REM" & " " & txtNick.Text
    DoEvents ' Makes SURE the data is being sent!
    End
Else
    End
End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 ' Index of the connection - button
        WS.Connect txtIP.Text, 1000 ' Connects to the server
        txtNick.Locked = True ' Locks the txtNick txtbox
End Select
End Sub
Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' if you press enter
    KeyAscii = 0 ' Resets..
    If Not txtSend.Text = "" Then ' If not the textbox is empty..
        'txtMotta.Text = txtSend.Text + vbCrLf
        WS.SendData " <" & txtNick.Text & ">" & " " & txtSend.Text + vbCrLf ' Sends the nick and the text in the send-box
        txtSend.Text = "" ' Makes the txtbox ready to send new text ( Empty )
    End If
End If
End Sub

Private Sub WS_Connect()
MsgBox "Connected to : " & WS.RemoteHostIP, vbInformation, "Connected" ' If you are connected, you get a msgbox
WS.SendData "ADD" & " " & txtNick.Text + vbCrLf ' Sends "ADD" when you connect, so server knows you are connected
End Sub
Function IsNickInList(Nick As String)
' This function will find what index the nick has
Dim I As Integer
IsNickInList = -1
For I = 0 To lstNicks.ListCount - 1
    If Nick = lstNicks.List(I) Then IsNickInList = I
Next
End Function
Function ErIliste(Nick As String) As Boolean
' This function will find out if the nick is in the nick-listbox
Dim j As Integer
For j = 0 To lstNicks.ListCount - 1
    If Nick = lstNicks.List(j) Then ErIliste = True
Next
End Function
Sub MotattLinje(S As String) ' This sub is used for all incoming data, check the dataarrival code :)
'Dim S As String
Dim Del() As String, SimilarNicksIndex As Integer, I As Integer, liste() As String
Del = Split(S, " ", 3) ' Splits up the data you receive.
If Del(0) = "REM" Then ' If the first word you receive is "REM"
    lstNicks.RemoveItem IsNickInList(Del(1))
    txtMotta.SelColor = vbRed
    txtMotta.SelText = "* " & Del(1) & " has disconnected" + vbCrLf
    txtMotta.SelColor = vbBlack
ElseIf Del(0) = "ADD" Then ' If the first word you receive is "ADD"
    lstNicks.AddItem Del(1)
    txtMotta.SelColor = vbBlue
    txtMotta.SelText = "* " & Del(1) & " has connected" + vbCrLf
    txtMotta.SelColor = vbBlack
ElseIf Del(0) = "LISTE" Then ' If the first word you receive is "LISTE"
    liste = Split(Del(2), " ")
    For I = 0 To UBound(liste)
        lstNicks.AddItem liste(I)
    Next
ElseIf Del(0) = "NICKINFO" Then ' If the first word you receive is "NICKINFO"
    txtMotta.SelColor = vbRed ' The seltext will be red then..
    txtMotta.SelText = "* Server modified your nick from: " & txtNick.Text & " to " & Del(1) & " due to >1 identic nicks" ' Text to put in the txtBox
    MsgBox "Server modified your nick from: " & txtNick.Text & " to " & Del(1), vbInformation, "NichChange"
    txtNick.Locked = False
    txtNick.Text = Del(1) ' the new nick will be put in the txtbox ready to be sent to server when sending text :)
    txtNick.Locked = True ' Locked..
Else
txtMotta.SelColor = vbBlack ' The seltext will then be black

txtMotta.SelColor = vbBlue ' The seltext will then be blue
txtMotta.SelText = Left(Del(1), 1) ' SelText = the leftmost character in Del(1)
txtMotta.SelColor = vbBlack ' Seltext will be black..
txtMotta.SelText = Mid(Del(1), 2, Len(Del(1)) - 2) ' the middle text from charachter two to the almost last character
txtMotta.SelColor = vbRed ' Seltext will then be red :)
txtMotta.SelText = Right(Del(1), 1) '
txtMotta.SelColor = vbBlack
txtMotta.SelText = " " & Del(2) + vbCrLf
'txtMotta.Text = txtMotta.Text + S + vbCrLf ' The chat-txtBox has now the text which the client received aswell
End If
End Sub
Private Sub WS_DataArrival(ByVal bytesTotal As Long)
'Dim S As String, Del() As String, SimilarNicksIndex As Integer, I As Integer, liste() As String
'WS.GetData S, vbString ' Gets the data being sent to clients from server

Dim S As String, buffer As String
Dim v As Integer
Dim C
WS.GetData S, vbString, bytesTotal
For v = 1 To Len(S) ' V = 1 to the lenght of S
C = Mid(S, v, 1) ' the midle of S, v , 1
If (C = vbCr) Or (C = vbLf) Then
If buffer <> "" Then MotattLinje buffer: buffer = "" ' runs the mottattlinje sub
Else
buffer = buffer + C
End If
Next
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, _
        ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, _
        ByVal HelpContext As Long, CancelDisplay As Boolean)
  MsgBox "Error: " & Description, vbCritical, "Error" ' Shows the error-message
  WS.SendData "REM" & " " & txtNick.Text + vbCrLf
End Sub

