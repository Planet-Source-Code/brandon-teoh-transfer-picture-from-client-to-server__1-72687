VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Client"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5700
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   16
      Text            =   "localhost"
      Top             =   240
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Select Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtPixPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1320
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3960
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send Text Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtPortNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Text            =   "1002"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Start Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4440
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Server IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Port Num:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sendImagePortNum As Integer
Private curPicture As StdPicture  'use StdPicture variable type
Private imageBytes() As Byte
Private serverIP As String


Private Sub Command1_Click()
   Call Form2.Show
End Sub

Private Sub Command2_Click()
 
If txtIP.Text = "" Then
   serverIP = "localhost"
Else
   serverIP = Trim(txtIP.Text)
End If
 
If Winsock1.State = 0 Then
    If Not txtPortNum.Text = "" Then
       Winsock1.RemotePort = CInt(txtPortNum.Text)
       Call Winsock1.Connect(serverIP, CInt(txtPortNum.Text))
    End If
Else
   MsgBox "To reconnect, close current connection!", vbCritical
End If
 
 
End Sub

Private Sub Command3_Click()
  Dim tempStr As String
  
 
  tempStr = Text1.Text
  If tempStr = "" Then
     tempStr = "Hello brandon!"
  End If
  
  If Winsock1.State = 7 Then
     Call Winsock1.SendData(tempStr)
  End If
  
  Text1.Text = ""
  
End Sub

Private Sub Command4_Click()
  Dim tempStr As String
  
  If Winsock1.State = 7 Then
   
     If Not Winsock2.State = 7 Then
        '==== Request to send image to server ===========
        tempStr = "Request_To_Send_Image"
  
         Call Winsock1.SendData(tempStr)
     
        '=================================================
     Else
        Call SendImageSize
     End If
  
  Else
     MsgBox "Server not connected!", vbCritical
  End If
  
End Sub

Private Sub Command5_Click()
 Call Winsock2.Close
 Call Winsock1.Close
 
 '=== Update labels===========================
 
 Label3.Caption = "Server NOT connected!"
 Label4.Caption = "Image socket NOT connected!"
 
End Sub

Private Sub Command6_Click()

Dim tempPath As String


CommonDialog1.CancelError = False 'this is important if the cancer is pressed
    'set default path
CommonDialog1.InitDir = App.Path
    ' Set flags
CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
CommonDialog1.Filter = "All Files (*.*)|*.*|JPG Files (*.jpg)|*.jpg"
    ' Specify default filter
CommonDialog1.FilterIndex = 2
    
CommonDialog1.ShowOpen  'display the dialog box with 'open' button

tempPath = CStr(CommonDialog1.FileName)

If Not tempPath = "" Then
   txtPixPath.Text = tempPath
   
   '=== Use LoadPicture (global method) ====
   
   Set curPicture = LoadPicture(txtPixPath.Text)
   
   If Not curPicture Is Nothing Then
      'Picture1.Picture = curPicture
      
      '=== resize picture
      Call StretchSourcePictureFromPicture(curPicture, Picture1)
      
   End If
   
End If
   
End Sub

Private Sub Winsock1_Connect()
'Dim tempNum As Integer

'tempNum = MsgBox("Server connected!", , "Client")
   
Label3.Caption = "Server connected!"
  
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim tempStr As String
 Dim tempNum As Integer
 
 Call Winsock1.GetData(tempStr)
  
 If Not tempStr = "" Then
    If Left(tempStr, 23) = "Send_Image_Port_Number=" Then
       sendImagePortNum = Right(tempStr, Len(tempStr) - 23)
       Call InitSendImage
    ElseIf tempStr = "Request_Image_Size" Then
       Call SendImageSize
    ElseIf tempStr = "Start_Send_Image" Then
       Call SendImage
    Else
      Text1.Text = Text1.Text & tempStr & vbCrLf
    End If
    
 End If
 
End Sub

Private Sub InitSendImage()
  Call StartImageSocket
 
End Sub

Private Sub StartImageSocket()
  
  If Not serverIP = "" Then
      If Not sendImagePortNum = 0 Then
         Winsock2.RemotePort = sendImagePortNum
         Call Winsock2.Connect(serverIP, sendImagePortNum)
      End If
  Else
     MsgBox "Unable to connect image socket", vbCritical
  End If
  
  End Sub

Private Sub SendImageSize()

Dim tempStr As String

If Not curPicture Is Nothing And Not txtPixPath.Text = "" Then
        Call ReadPictureIntoBytes
        
        If Winsock1.State = 7 Then
           
           Label2.Caption = "Total bytes to send=" & UBound(imageBytes)
           
           If UBound(imageBytes) > -1 Then
              tempStr = "Image_Size=" & (UBound(imageBytes) + 1)
              Call Winsock1.SendData(tempStr)
           End If
           
        End If
 Else
        MsgBox "No image is selected!"
        '=== Close image socket ====
        Winsock2.Close
 End If
    
End Sub

Private Sub SendImage()
    
    If Not curPicture Is Nothing And Not txtPixPath.Text = "" Then
        Call ReadPictureIntoBytes
        
        '=== For debugging =====================
        'Call CopyImageContent2File
        
        If Winsock2.State = 7 Then
           
           If UBound(imageBytes) > -1 Then
              Call Winsock2.SendData(imageBytes)
           End If
           
        End If
    Else
        MsgBox "No image is selected!"
        '=== Close image socket ====
        Winsock2.Close
    End If
    
End Sub

Private Sub Winsock2_Connect()
   'Dim tempNum As Integer

   'tempNum = MsgBox("Image socket connected!", , "Client")
   Label4.Caption = "Image socket connected!"
   
End Sub

Private Sub ReadPictureIntoBytes()

  Dim iFileNum As Integer
  Dim lFileLength As Long
  Dim sTempFile As String
 
  sTempFile = txtPixPath.Text

    'read file contents to byte array
    iFileNum = FreeFile
    Open sTempFile For Binary Access Read As #iFileNum
    lFileLength = LOF(iFileNum)
    ReDim imageBytes(lFileLength)
    Get #iFileNum, , imageBytes()
    
    Close #iFileNum
    
End Sub

Private Sub CopyImageContent2File(Optional sTempFile As String)

 Dim i As Long
 Dim tempASCIINum As Integer
 Dim tempStr As String
 Dim fso As New FileSystemObject
 
 If UBound(imageBytes) > -1 Then
                
    If sTempFile = "" Then
       sTempFile = App.Path & "\Data" & "\" & "ClientData.txt"
    End If
    
    '=== delete the temp file away if exists ====
     If fso.FileExists(sTempFile) Then
       Kill sTempFile
    End If
    
    For i = 0 To UBound(imageBytes)
        tempASCIINum = Asc(imageBytes(i))
        tempStr = "i=" & i & ";ASCII=" & tempASCIINum & vbCrLf
        Call writeFileData(tempStr, "Append", sTempFile)
        Label5.Caption = "Copying image content i=" & i
        DoEvents
    Next i
               
 End If

 Set fso = Nothing
 
End Sub
