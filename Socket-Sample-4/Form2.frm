VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Server"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6795
   LinkTopic       =   "Form2"
   ScaleHeight     =   3975
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   4680
      ScaleHeight     =   2115
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   600
      Width           =   3975
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   6240
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPortNum 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Text            =   "1002"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Listening"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   5295
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   5295
   End
   Begin VB.Label lblStatus 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Port Num:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private imagePortNum As Integer
Private imageByte() As Byte
Private expectedImageSize As Long
Private curPicture As StdPicture  'use StdPicture variable type
Private curImageSize As Long
Private imageByteCol As Collection
Private imageObj As Variant

Private Sub RefreshImageByteCol()
 Set imageByteCol = Nothing
 Set imageByteCol = New Collection
End Sub

Private Sub Command1_Click()

Call InitDataSocket

End Sub

Private Sub InitDataSocket()
If Not txtPortNum.Text = "" Then
   Winsock1.LocalPort = CInt(txtPortNum.Text)
   Call Winsock1.Listen
   lblStatus.Caption = "Text socket state=" & Winsock1.State
End If
End Sub

Private Sub Form_Load()
 
 '===ReDim it temporarily first ====
 
 ReDim imageByte(0)
 
End Sub



Private Sub Winsock1_Close()
  Winsock1.Close
  lblStatus.Caption = "Text socket state=" & Winsock1.State
  
  '=== init again ========
  Call InitDataSocket
  
  '==== refresh form ======
  Call RefreshForm
  
End Sub

Private Sub Winsock1_Connect()
   Dim tempNum As Integer
   
   tempNum = MsgBox("Client connected!", , "Server")
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
  Call Winsock1.Close
  Call Winsock1.Accept(requestID)
  lblStatus.Caption = "Text socket state=" & Winsock1.State
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
 Dim tempStr As String
 Dim tempNum As Integer
 
 Call Winsock1.GetData(tempStr)
  
 If Not tempStr = "" Then
    If tempStr = "Request_To_Send_Image" Then
       Call InitReceiveImage
    ElseIf Left(tempStr, 11) = "Image_Size=" Then
       expectedImageSize = CLng(Right(tempStr, Len(tempStr) - 11))
       If Not expectedImageSize = 0 Then
          Call InitProgressBar
          Call refreshVariables
          Call NotifyClientToSendImage
       End If
    Else
      'tempNum = MsgBox(tempStr, , "Server")
      Text1.Text = Text1.Text & tempStr & vbCrLf
    End If
    
 End If
 
End Sub

Private Sub refreshVariables()
 
 curImageSize = 0
 ReDim imageByte(0)
 imageObj = ""
 Set curPicture = Nothing
 
End Sub

Private Sub InitReceiveImage()

   Call StartImageSocket
   Call NotifyClientAbtImageSocket
End Sub

Private Sub StartImageSocket()
   imagePortNum = 1003
   Winsock2.LocalPort = imagePortNum
   Call Winsock2.Listen
   Label2.Caption = "Image socket state=" & Winsock2.State
End Sub

Private Sub NotifyClientAbtImageSocket()
  Dim tempStr As String
  
  tempStr = "Send_Image_Port_Number=" & imagePortNum
  
  If Winsock1.State = 7 Then
     Call Winsock1.SendData(tempStr)
  End If
  
End Sub

Private Sub NotifyClientToSendImage()
  Dim tempStr As String
  
  tempStr = "Start_Send_Image"
  
  If Winsock1.State = 7 Then
     Call Winsock1.SendData(tempStr)
  End If
  
  '=== refresh collection ==========
  Call RefreshImageByteCol
  
End Sub

Private Sub RequestForImageSize()
  Dim tempStr As String
  
  tempStr = "Request_Image_Size"
  
  If Winsock1.State = 7 Then
     Call Winsock1.SendData(tempStr)
  End If
  
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
  Call Winsock2.Close
  Call Winsock2.Accept(requestID)
  Label2.Caption = "Image socket state=" & Winsock2.State
  'Call NotifyClientToSendImage
  Call RequestForImageSize
 
End Sub

Private Sub Winsock2_Connect()
    Dim tempNum As Integer
   
   tempNum = MsgBox("Image socket connected!", , "Server")
 
End Sub


Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
   
   '==================================================
   'curImageSize = curImageSize + bytesTotal  'This works as well
   curImageSize = curImageSize + Winsock2.BytesReceived
   
   '==== Method 4 (This works) =================
   Dim tempImageObj As Variant
   Call Winsock2.GetData(tempImageObj)
   Call ConsolidateObj(tempImageObj)
   '============================================
   
End Sub

Private Sub Winsock2_Close()
  Winsock2.Close
  Label2.Caption = "Image socket state=" & Winsock2.State
End Sub


Private Sub ConsolidateObj(tempData As Variant)

'=== consolidate data into variant variable ====
imageObj = imageObj & tempData

'=== update progress bar (this is risky, it may cause winsock data buffer to corrupt if variant doesn't employ thread concurrency =======
Call UpdateProgressBar(curImageSize)

If curImageSize >= expectedImageSize Then
    'MsgBox curImageSize
    
    Call ConsolidateObjSubProcess2
    
    Call DisplayImage
    
    'MsgBox "Send complete"
        
End If
 
End Sub

'=== This works =========================
Private Sub ConsolidateObjSubProcess2()
  '=== Assign variant to byte array====
  imageByte = imageObj
  
End Sub

Private Sub ConsolidateArraySubProcess3()

Dim i As Integer
Dim k As Long
Dim tempArray() As Byte
Dim tempUBound As Long
Dim lastUBound As Long
Dim newUBound As Long
Dim startIndex As Long
Dim totalUBound As Long
Dim totalUBound2 As Long

If Not imageByteCol Is Nothing Then
 
   If Not imageByteCol.Count = 0 Then
      
      '===== Perform redim first ==================
       For i = 1 To imageByteCol.Count
          tempArray = imageByteCol.Item(i)
          tempUBound = UBound(tempArray)
          
          If i = 1 Then 'First time
             totalUBound = tempUBound
          Else
             totalUBound = totalUBound + tempUBound + 1
          End If
          
       Next i
       
       ReDim imageByte(totalUBound)
       
     '===============================================
      
     '===== Combine items ===========================
      For i = 1 To imageByteCol.Count
          tempArray = imageByteCol.Item(i)
          tempUBound = UBound(tempArray)
          
          If lastUBound = 0 Then
             newUBound = tempUBound
             startIndex = 0
             totalUBound2 = tempUBound
          Else
             newUBound = lastUBound + tempUBound + 1
             startIndex = lastUBound + 1
             totalUBound2 = totalUBound2 + tempUBound + 1
          End If
            
          Call ConsolidateArraySubProcess2_2(startIndex, newUBound, tempArray)
          
          lastUBound = tempUBound
          
          '=== update progress bar =======
          Call UpdateProgressBar(totalUBound2)

       Next i
      '=============================================
   
   End If ' If Not imageByteCol.Count = 0 Then

End If 'If Not imageByteCol Is Nothing Then

End Sub

Private Sub ConsolidateArraySubProcess2_2(startIndex As Long, newUBound As Long, tempImageByte() As Byte)

    Dim i As Long
    
    For i = startIndex To newUBound
        imageByte(i) = tempImageByte(i - startIndex)
    Next i
    
End Sub

Private Sub DisplayImage()

Dim iFileNum As Integer
Dim lFileLength As Long
Dim abBytes() As Byte
Dim sTempFile As String
 
      '==== For debugging ==================
      'Call CopyImageContent2File
      '=====================================
      
      sTempFile = App.Path & "\pix\" & Format(Now, "yyyymmddhhnnss") & ".jpg"
  
      iFileNum = FreeFile
      Open sTempFile For Binary As #iFileNum
      Put #iFileNum, , imageByte()
      Close #iFileNum
      
      If Not sTempFile = "" Then
         Set curPicture = LoadPicture(sTempFile)
   
         If Not curPicture Is Nothing Then
            'Picture1.Picture = curPicture
            
              '=== resize picture
              Call StretchSourcePictureFromPicture(curPicture, Picture1)
      
         End If
      End If
      
       '=== delete the temp file away ====
      Kill sTempFile
End Sub

Private Sub InitProgressBar()

ProgressBar1.Value = 0
DoEvents
ProgressBar1.Min = 0
ProgressBar1.Max = expectedImageSize + 1

End Sub

Private Sub UpdateProgressBar(curValue As Long)

ProgressBar1.Value = curValue
DoEvents

End Sub

Private Sub RefreshForm()

Text1.Text = ""
Set Picture1.Picture = Nothing

End Sub

Private Sub CopyImageContent2File(Optional sTempFile As String)

 Dim i As Long
 Dim tempASCIINum As Integer
 Dim tempStr As String
 Dim fso As New FileSystemObject
 
 If UBound(imageByte) > -1 Then
                
    If sTempFile = "" Then
       sTempFile = App.Path & "\Data" & "\" & "ServerData.txt"
    End If
    
    '=== delete the temp file away if exists ====
    If fso.FileExists(sTempFile) Then
       Kill sTempFile
    End If
  
    For i = 0 To UBound(imageByte)
        tempASCIINum = Asc(imageByte(i))
        tempStr = "i=" & i & ";ASCII=" & tempASCIINum & vbCrLf
        Call writeFileData(tempStr, "Append", sTempFile)
        Label3.Caption = "Copying image content i=" & i
        'DoEvents
    Next i
               
 End If

 Set fso = Nothing
 
End Sub

Private Sub CopyImageContent2File2(dataByte() As Byte, Optional sTempFile As String)

 Dim i As Long
 Dim tempASCIINum As Integer
 Dim tempStr As String
 Dim fso As New FileSystemObject
 
 If UBound(dataByte) > -1 Then
                
    If sTempFile = "" Then
       sTempFile = App.Path & "\Data" & "\" & "ServerData.txt"
    End If
    
    '=== delete the temp file away if exists ====
    If fso.FileExists(sTempFile) Then
       Kill sTempFile
    End If
  
    For i = 0 To UBound(dataByte)
        tempASCIINum = Asc(dataByte(i))
        tempStr = "i=" & i & ";ASCII=" & tempASCIINum & vbCrLf
        Call writeFileData(tempStr, "Append", sTempFile)
        Label3.Caption = "Copying image content i=" & i
        'DoEvents
    Next i
               
 End If

 Set fso = Nothing
 
End Sub
