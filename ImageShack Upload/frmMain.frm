VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ImageShack Uploader"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   6960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   7440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CheckBox chkRandomName 
      Caption         =   "Give online image a random name"
      Height          =   255
      Left            =   3840
      TabIndex        =   34
      Top             =   4600
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   8040
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCopyAll 
      Caption         =   "Copy All"
      Height          =   315
      Left            =   7200
      TabIndex        =   31
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdInet 
      Caption         =   "Upload (Inet)"
      Height          =   315
      Left            =   2040
      TabIndex        =   30
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdWinsock 
      Caption         =   "Upload (Winsock)"
      Height          =   315
      Left            =   120
      TabIndex        =   29
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Upload Results"
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   8295
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   180
         ScaleHeight     =   2895
         ScaleWidth      =   7995
         TabIndex        =   4
         Top             =   360
         Width           =   8000
         Begin VB.TextBox txtResult 
            Height          =   285
            Index           =   7
            Left            =   720
            TabIndex        =   20
            Top             =   2520
            Width           =   5295
         End
         Begin VB.TextBox txtResult 
            Height          =   285
            Index           =   6
            Left            =   720
            TabIndex        =   19
            Top             =   2160
            Width           =   5295
         End
         Begin VB.TextBox txtResult 
            Height          =   285
            Index           =   5
            Left            =   720
            TabIndex        =   18
            Top             =   1800
            Width           =   5295
         End
         Begin VB.TextBox txtResult 
            Height          =   285
            Index           =   4
            Left            =   720
            TabIndex        =   17
            Top             =   1440
            Width           =   5295
         End
         Begin VB.TextBox txtResult 
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   16
            Top             =   1080
            Width           =   5295
         End
         Begin VB.TextBox txtResult 
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   15
            Top             =   720
            Width           =   5295
         End
         Begin VB.TextBox txtResult 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   14
            Top             =   360
            Width           =   5295
         End
         Begin VB.TextBox txtResult 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   13
            Top             =   0
            Width           =   5295
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   12
            Top             =   2540
            Width           =   615
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   11
            Top             =   2180
            Width           =   615
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   10
            Top             =   1820
            Width           =   615
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   9
            Top             =   1460
            Width           =   615
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   8
            Top             =   1100
            Width           =   615
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   7
            Top             =   740
            Width           =   615
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   6
            Top             =   380
            Width           =   615
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   20
            Width           =   615
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Direct link to image"
            Height          =   195
            Index           =   7
            Left            =   6120
            TabIndex        =   28
            Top             =   2540
            Width           =   1350
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Show image to friends"
            Height          =   195
            Index           =   6
            Left            =   6120
            TabIndex        =   27
            Top             =   2200
            Width           =   1560
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Hotlink for Websites"
            Height          =   195
            Index           =   5
            Left            =   6120
            TabIndex        =   26
            Top             =   1840
            Width           =   1425
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Hotlink for forums (2)"
            Height          =   195
            Index           =   4
            Left            =   6120
            TabIndex        =   25
            Top             =   1480
            Width           =   1455
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Hotlink for forums (1)"
            Height          =   195
            Index           =   3
            Left            =   6120
            TabIndex        =   24
            Top             =   1120
            Width           =   1455
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Thumbnail for forums (2)"
            Height          =   195
            Index           =   2
            Left            =   6120
            TabIndex        =   23
            Top             =   760
            Width           =   1695
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Thumbnail for forums (1)"
            Height          =   195
            Index           =   1
            Left            =   6120
            TabIndex        =   22
            Top             =   400
            Width           =   1695
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Thumbnail for Websites"
            Height          =   195
            Index           =   0
            Left            =   6120
            TabIndex        =   21
            Top             =   40
            Width           =   1665
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Image"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   1065
         TabIndex        =   32
         Top             =   340
         Width           =   1065
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse"
            Height          =   315
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   6915
         TabIndex        =   1
         Top             =   370
         Width           =   6915
         Begin VB.TextBox txtImagePath 
            Height          =   285
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   6855
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'// Note: The uploads might take a little while, depending
'// on the image size and how busy the ImageShack server is.

'// This is only an example how to upload an image with Inet and Winsock.
'// If you want to upload multiple images at once, then you'll have to create a Winsock array yourself.


Option Explicit
Option Compare Text

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private WinsockHTML As String '// Only used when uploading with Winsock to hold the html source code

'// Browse for an image
Private Sub cmdBrowse_Click()
On Error GoTo LastLine
    
    With CDlg
        .Flags = cdlOFNExplorer + cdlOFNHideReadOnly
        .Filter = "Image Files (*.jpg;*.jpeg;*.png;*.gif;*.bmp)|*.jpg;*.jpeg;*.png;*.gif;*.bmp"
        .DialogTitle = "Select an image file"
        .CancelError = True
        .ShowOpen
        If .FileName <> "" Then
            If FileLen(.FileName) >= 3145728 Then '// 3mb max
                MsgBox "This image is too big. Maximum file size is 3mb. ", vbExclamation, "Image too big"
                Exit Sub
            Else
                txtImagePath.Text = .FileName
            End If
        End If
    End With

LastLine:
End Sub

'// Copy a single image link to the Clipboard
Private Sub cmdCopy_Click(Index As Integer)
    Clipboard.Clear
    Clipboard.SetText txtResult(Index).Text
End Sub

'// Copy all image links to the Clipboard at once
Private Sub cmdCopyAll_Click()
    Dim i As Integer
    Dim tmp As String
    
    For i = 0 To 7
        If txtResult(i).Text <> "n.a" Then
            tmp = tmp & txtResult(i).Text & vbCrLf
        End If
    Next i
    
    Clipboard.Clear
    Clipboard.SetText tmp
End Sub

'// Upload with Inet ====================================================
Private Sub cmdInet_Click()
    Dim arr() As String
    
    On Error GoTo ErrHandler
    
    If chkRandomName.Value Then
        arr = PrepareImageUpload(txtImagePath.Text, m_Inet, True) '// Create random image name
    Else
        arr = PrepareImageUpload(txtImagePath.Text, m_Inet) '// Keep original image name
    End If
    
    DisableResults
    DisableButtons
    Me.Caption = "ImageShack Uploader - Uploading. Please wait..."
    
    Inet.RequestTimeout = 10
    Inet.Execute "http://www.imageshack.us", "POST", arr(0), arr(1) '// arr(0) = Body and arr(1) = Header
    
Exit Sub
ErrHandler:
    EnableButtons
    Me.Caption = "ImageShack Uploader"
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub Inet_StateChanged(ByVal State As Integer)
   Dim vtData   As Variant
   Dim strData  As String
   Dim bDone    As Boolean
   Dim arr()    As String
   Dim i        As Integer
   
   Select Case State
   Case icError ' 11
      '// In case of error, return ResponseCode and ResponseInfo.
      vtData = Inet.ResponseCode & " - " & Inet.ResponseInfo
   Case icResponseCompleted ' 12
      bDone = False
      '// Get first chunk.
      vtData = Inet.GetChunk(1024, icString)
      DoEvents
      Do While Not bDone
         strData = strData & vtData
         '// Get next chunk.
         vtData = Inet.GetChunk(1024, icString)
         DoEvents
         If Len(vtData) = 0 Then
            bDone = True
         End If
      Loop
      '// Grab the links from the html source code
      arr = GrabLinks(strData)
      
      EnableResults
      EnableButtons
      
      '// Show links in the textboxes
      For i = 0 To 7
        txtResult(i).Text = arr(i)
      Next i
      Me.Caption = "ImageShack Uploader"
   End Select
   
End Sub

'// Upload with Winsock =================================================
Private Sub cmdWinsock_Click()
    
    On Error GoTo ErrHandler
    
    DisableResults
    DisableButtons
    Me.Caption = "ImageShack Uploader - Uploading. Please wait..."
    
    Winsock.Close
    Winsock.Connect "imageshack.us", 80
    
Exit Sub
ErrHandler:
    EnableButtons
    Me.Caption = "ImageShack Uploader"
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub Winsock_Connect()
    Dim arr() As String
    
    If chkRandomName.Value Then
        arr = PrepareImageUpload(txtImagePath.Text, m_Winsock, True) '// Create random image name
    Else
        arr = PrepareImageUpload(txtImagePath.Text, m_Winsock) '// Keep original image name
    End If
    
    Winsock.SendData arr(0) '// arr(0) = Header + Body in one piece
End Sub

Private Sub Winsock_DataArrival(ByVal BytesTotal As Long)
    Dim sData   As String
    Dim arr()   As String
    Dim i       As Integer
    
    Winsock.GetData sData, vbString
    WinsockHTML = WinsockHTML & sData
    
    '// If entire html page has been returned
    If InStr(WinsockHTML, "</html>") Then
        '// Grab the links from the html source code
        arr = GrabLinks(WinsockHTML)
      
        EnableResults
        EnableButtons
        
        '// Show links in the textboxes
        For i = 0 To 7
            txtResult(i).Text = arr(i)
        Next i
        Me.Caption = "ImageShack Uploader"
    
        WinsockHTML = vbNullString
    End If
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Winsock.Close
    MsgBox Number & " - " & Description
End Sub

Private Sub Form_Load()
    DisableResults
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Terminate()
    Winsock.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock.Close
End Sub

Private Sub EnableResults()
    Dim i As Integer
    
    For i = 0 To 7
        cmdCopy(i).Enabled = True
        txtResult(i).Enabled = True
        txtResult(i).Text = vbNullString
        txtResult(i).BackColor = vbWhite
        Label(i).Enabled = True
    Next i
    cmdCopyAll.Enabled = True
End Sub

Private Sub EnableButtons()
    cmdBrowse.Enabled = True
    cmdWinsock.Enabled = True
    cmdInet.Enabled = True
    chkRandomName.Enabled = True
End Sub

Private Sub DisableResults()
    Dim i As Integer
    
    For i = 0 To 7
        cmdCopy(i).Enabled = False
        txtResult(i).Enabled = False
        txtResult(i).Text = vbNullString
        txtResult(i).BackColor = vbButtonFace
        Label(i).Enabled = False
    Next i
    cmdCopyAll.Enabled = False
End Sub

Private Sub DisableButtons()
    cmdBrowse.Enabled = False
    cmdWinsock.Enabled = False
    cmdInet.Enabled = False
    chkRandomName.Enabled = False
End Sub
