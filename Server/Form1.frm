VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Test"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtMotta 
      Height          =   2355
      Left            =   0
      TabIndex        =   6
      Top             =   660
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   4154
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.ListBox lstNicks 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   2400
      TabIndex        =   5
      Top             =   660
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   1380
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
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   1111
      ButtonWidth     =   926
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Listen"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.TextBox txtNick 
         Height          =   285
         Left            =   660
         TabIndex        =   4
         Text            =   "Server"
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   315
      Left            =   2460
      TabIndex        =   2
      Top             =   3060
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   3465
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   344
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "00:05"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   3060
      Width           =   2355
   End
   Begin MSWinsockLib.Winsock WS 
      Index           =   0
      Left            =   3540
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ClientIndex As Integer
' This code can be used in own, or any project as you'd like.. I have no "copyright" in it
' because I like to share my code :) Enjoy, and PLEASE go back and vote for it, i really need it.
Private Sub cmdSend_Click()
Dim i As Integer
If Not txtSend.Text = "" Then ' If empty, then..
For i = 1 To ClientIndex - 1 ' All of the connected clients
    txtMotta.Text = txtMotta.Text + " <" & txtNick.Text & ">" & txtSend.Text + vbCrLf
    WS(i).SendData " <" & txtNick.Text & ">" & " " & txtSend.Text + vbCrLf ' sends to all
Next
End If
End Sub

Private Sub Form_Load()
ClientIndex = 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        WS(0).LocalPort = 1000
        WS(0).Listen
End Select
End Sub

'Private Sub WS_Close(Index As Integer)
'ClientIndex = ClientIndex - 1
'WS(Index).Close
'Unload WS(Index)
'End Sub

Private Sub WS_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim Parts() As String, S As String, NickName As String, i As Integer

If ClientIndex = 1 Then
    Load WS(ClientIndex) ' Loads another sock
    WS(ClientIndex).Close ' Closes it before using it
    WS(ClientIndex).Accept requestID ' Accepts the incoming connection
    ClientIndex = ClientIndex + 1 ' Clientindex will here be 2
    
ElseIf ClientIndex > 1 Then ' If ClientIndex is higher than 1 then
    Load WS(ClientIndex) ' Loads a new socket
    WS(ClientIndex).Close ' Closes it before using it
    WS(ClientIndex).Accept requestID ' Accepts the incomming connection
    ClientIndex = ClientIndex + 1 ' Adds 1 to clientindex
End If
StatusBar1.Panels(2).Text = "Clients connected: " & ClientIndex - 1 ' shows the ammount of
                                                                    ' connected clients in
                                                                    ' the statysbar :)
End Sub
Function IsNickInList(Nick As String)
' This function will find out what index "Nick" has in the nicklistbox
Dim i As Integer
IsNickInList = -1
For i = 0 To lstNicks.ListCount - 1
    If Nick = lstNicks.List(i) Then IsNickInList = i
Next
End Function
Function ErIliste(Nick As String) As Boolean
' This function will find out if "Nick" is in the nicklistbox or not
Dim j As Integer
For j = 0 To lstNicks.ListCount - 1
    If Nick = lstNicks.List(j) Then ErIliste = True
Next
End Function
Sub MotattLinje(S As String, Index As Integer)
On Error Resume Next

Dim i As Integer, sck As Object
Dim Del() As String, SimilarNicksIndex  As Integer, test As String, NickName As String

'WS(Index).GetData S, vbString
Del = Split(S, " ", 2)
If Del(0) = "ADD" Then
    If ErIliste(Del(1)) Then ' If nick is in the list then..
        SimilarNicksIndex = 1
        While ErIliste(Del(1) & "(" & SimilarNicksIndex & ")")
            SimilarNicksIndex = SimilarNicksIndex + 1
       Wend
        lstNicks.AddItem Del(1) + "(" & SimilarNicksIndex & ")" ' adds the nick
'        For i = 1 To ClientIndex - 1
'            WS(i).SendData Del(1) & "(" & SimilarNicksIndex & ")"
'        Next
            WS(Index).SendData "NICKINFO" & " " & Del(1) & "(" & SimilarNicksIndex & ")" & " " + vbCrLf
            NickName = Del(1) & "(" & SimilarNicksIndex & ")"
    Else
        lstNicks.AddItem Del(1)
        NickName = Del(1)
    End If
    
    For i = 0 To lstNicks.ListCount - 1
        test = test & " " & lstNicks.List(i)
        'test = test & " " & lstNicks.Text
    Next
    
    'WS(ClientIndex - 1).SendData "LISTE" & " " & test + vbCrLf
    WS(Index).SendData "LISTE " & test + vbCrLf
    txtMotta.SelColor = vbBlue
    txtMotta.SelText = "* " & Del(1) & " has connected"
    
    For i = 1 To ClientIndex - 1
        If i <> Index Then
            'WS(i).SendData "ADD" & " " & Del(1) + vbCrLf
            WS(i).SendData "ADD" & " " & NickName + vbCrLf
        Else
        End If
    Next
ElseIf Del(0) = "REM" Then

    lstNicks.RemoveItem IsNickInList(Del(1))
    txtMotta.SelColor = vbRed
    txtMotta.SelText = "* " & Del(1) & " has disconnected" + vbCrLf

'    For i = 1 To ClientIndex   ' ALL the clients
    For i = 1 To ClientIndex - 1
        WS(i).SendData "REM" & " " & Del(1) + vbCrLf ' sends "REM" & the nickname which is disconnecting
    Next
    
    'ClientIndex = ClientIndex - 1
    WS(ClientIndex).Close  ' closes the used sock
    Unload WS(ClientIndex) ' Unloads the socket used ( WIll make a hole in the list, but
                           ' no problem.. since the ammount of listening socks/ used can
                           ' be up to.. 32000
Else
    For i = 1 To ClientIndex - 1
        WS(i).SendData S + vbCrLf ' ws(i)... sends S and new line to make sure there wont be
                                  ' two commands sent at once (Had some trouble liek that)
    Next
    txtMotta.SelColor = vbBlack ' selcolor is then black
    txtMotta.SelText = Del(1) + vbCrLf ' seltext = del(1)
End If
End Sub

Private Sub WS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim S As String, Buffer As String
Dim v As Integer
Dim C
WS(Index).GetData S, vbString, bytesTotal ' Gets the data as vbstring, bytestotal
For v = 1 To Len(S)
C = Mid(S, v, 1)
If (C = vbCr) Or (C = vbLf) Then
If Buffer <> "" Then MotattLinje Buffer, Index: Buffer = ""
Else
Buffer = Buffer + C
End If
Next
End Sub

Private Sub WS_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print Index; " "; Number; " "; Description
WS(Index).Close

'MsgBox "Error: " & Description, vbCritical, "Error"
End Sub
