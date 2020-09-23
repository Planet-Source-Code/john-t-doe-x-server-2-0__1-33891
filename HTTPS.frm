VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "X-Server v2.0"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Stats 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4471
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Control"
      TabPicture(0)   =   "HTTPS.frx":0000
      Tab(0).ControlCount=   8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "IPL"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "IPA"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "IndexFile"
      Tab(0).Control(7).Enabled=   0   'False
      TabCaption(1)   =   "Log"
      TabPicture(1)   =   "HTTPS.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List1"
      Tab(1).Control(0).Enabled=   -1  'True
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox IndexFile 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "Index.html"
         Top             =   1560
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Block"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox IPA 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.ListBox IPL 
         Height          =   645
         Left            =   3240
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Stop"
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   2160
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Index File:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "IP Address"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1320
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hitnum As String

Private Sub Command1_Click()
IPL.AddItem IPA.Text
End Sub

Private Sub Form_Load()
Winsock1.Close
Winsock1.LocalPort = 80
Winsock1.Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End
End Sub

Private Sub Option1_Click()
Winsock1.Close
Winsock1.LocalPort = 80
Winsock1.Listen
End Sub

Private Sub Option2_Click()
Winsock1.Close
Winsock1.LocalPort = 10254
Winsock1.Listen
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
MsgBox "Connected"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
List1.AddItem "Connecting"
Winsock1.Accept requestID
For i% = 0 To IPL.ListCount
If Winsock1.RemoteHostIP = IPL.List(i%) Then
Winsock1.Close
Exit Sub
End If
Next
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim data As String 'Variable To Hold File Contents
Dim datax As String 'Variable For Incoming Data
Winsock1.GetData datax 'Retrieve Incoming Data
request = Mid(datax, 5, InStr(5, datax, " ") - 5) 'Gets requested filename from clients data
If request = "/" Then request = IndexFile.Text  'If no requested filename (EX_ www.funbrain.com has no selected filename)  direct user to Index.html file
Open App.Path + "\" + Mid(request, 2, Len(request) - 1) For Input As #1
Input #1, rec1
Close

Open App.Path + "\" + Mid(request, 2, Len(request) - 1) For Binary Access Read As #1 'Open requested file
If Err.Number Then 'See if there is an error
    Open App.Path + "\Error.html" For Output As FreeFile 'If an error has occured create an ERROR WEBPAGE
        Write #FreeFile - 1, "<HTML>"
        Write #FreeFile - 1, "<Body>"
        Write #FreeFile - 1, "ERROR!  PAGE DOESNT EXCIST"
        Write #FreeFile - 1, "</Body>"
        Write #FreeFile - 1, "</HTML>"
    Close
    Open App.Path + "\Error.html" For Binary Access Read As FreeFile 'Retrieve contents of error webpage
        data = Space(LOF(1))
        Get #FreeFile - 1, , data
    Close
    Else
        data = Space(LOF(1)) '
        Get #1, , data 'Retrieve file contents
End If
Close ' Close File to free up MEM

codeType = "text/html" 'Tell client what type of data is being sent
Winsock1.SendData "HTTP/1.0 200 OK" & vbCrLf & _
                                     "Content-Length: " & Len(data) & vbCrLf & _
                                     "Content-Type: " & codeType & vbCrLf & _
                                     vbCrLf & _
                                     data 'Send page with header
                                     List1.AddItem "DONE"
                                     List1.AddItem data
                                     DoEvents
Winsock1.Close
Winsock1.LocalPort = 80
Winsock1.Listen
End Sub

