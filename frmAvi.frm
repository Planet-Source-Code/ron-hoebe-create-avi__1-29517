VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAVI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create AVI from Picture Files"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmAvi.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   5460
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   60
      TabIndex        =   10
      Top             =   7440
      Width           =   1400
   End
   Begin VB.Frame Frame3 
      Caption         =   "Image"
      Height          =   5625
      Left            =   60
      TabIndex        =   9
      Top             =   1740
      Width           =   5325
      Begin VB.PictureBox imShow 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         Height          =   5000
         Left            =   180
         ScaleHeight     =   4935
         ScaleWidth      =   4935
         TabIndex        =   18
         Top             =   540
         Width           =   5000
      End
      Begin VB.PictureBox imL 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Height          =   675
         Left            =   4890
         ScaleHeight     =   615
         ScaleWidth      =   585
         TabIndex        =   17
         Top             =   5040
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox chkShowImage 
         Caption         =   "Show Image when creating AVI"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Value           =   1  'Checked
         Width           =   2745
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1095
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   5325
      Begin VB.CommandButton cmdSrc 
         Caption         =   "..."
         Height          =   315
         Left            =   4650
         TabIndex        =   13
         Top             =   240
         Width           =   525
      End
      Begin VB.CommandButton cmdDest 
         Caption         =   "..."
         Height          =   315
         Left            =   4650
         TabIndex        =   11
         Top             =   660
         Width           =   525
      End
      Begin VB.TextBox txtFPS 
         Height          =   285
         Left            =   660
         TabIndex        =   6
         Text            =   "10"
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Destination:"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Source:"
         Height          =   255
         Left            =   1530
         TabIndex        =   15
         Top             =   300
         Width           =   765
      End
      Begin VB.Label lblFileSrc 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Picture File -> Browse ->"
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   270
         Width           =   2175
      End
      Begin VB.Label lblFileDest 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AVI File -> Browse ->"
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   660
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "FPS:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   435
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2310
      Top             =   7410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Create AVI"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3990
      TabIndex        =   3
      Top             =   7440
      Width           =   1400
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   7890
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Progress"
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   1170
      Width           =   5325
      Begin VB.PictureBox Picture1 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   5025
         TabIndex        =   1
         Top             =   240
         Width           =   5085
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3990
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   1400
   End
End
Attribute VB_Name = "frmAVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cBar As ClsProgressBar2
Dim bmp As cDIB
Dim fPath As String
Dim szOutputAVIFile As String
Dim szInputFile As String
Dim docancel As Boolean
Dim telDir As Long
Dim fDest As Boolean
Dim fSrc As Boolean

Private Sub cmdCancel_Click()
    docancel = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDest_Click()
    fDest = DestFile
    If fSrc And fDest Then cmdGo.Enabled = True
End Sub

Private Sub cmdGo_Click()
    Dim msgString As String
    Dim bmpFile As String
    Dim sFile As String
    Dim res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim BI As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim i As Long

    docancel = False
    cmdGo.Visible = False
    cmdCancel.Visible = True
    cmdClose.Enabled = False
    cmdSrc.Enabled = False
    cmdDest.Enabled = False
    
'    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set bmp = New cDIB
    
    sFile = Dir$(fPath + "\*.*")
    bmpFile = loadPic(szInputFile, False, False, True)
    If bmp.CreateFromFile(bmpFile) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
        GoTo error
    End If

'   Fill in the header for the video stream
    If Val(txtFPS) < 1 Or Val(txtFPS) > 50 Then txtFPS = "10"
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&                             '// default AVI handler
        .dwScale = 1
        .dwRate = Val(txtFPS)                         '// fps
        .dwSuggestedBufferSize = bmp.SizeImage       '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)       '// rectangle for stream
    End With
    
    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

'   And create the stream
    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error

    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    res = AVISaveOptions(frmAVI.hWnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)
    'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then 'In C TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo error
    End If
    
    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error
    
    'set format of stream according to the bitmap
    With BI
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

'   Now write out each video frame
    cBar.Value = 1
    cBar.StartTimer
    DoEvents
    i = 0
    Do While sFile <> ""
        i = i + 1
        cBar.Value = i
        frmAVI.Frame1.Caption = "Estimated Time Remaining " & cBar.Time2End
        cBar.ShowBar
        DoEvents
        frmAVI.sBar.SimpleText = "Adding " + GetFileBaseName(sFile) + "." + GetFileExtension(sFile)
        bmpFile = loadPic(fPath + "\" + sFile, chkShowImage = vbChecked, False, True)
        DoEvents
        bmp.CreateFromFile (bmpFile) 'load the bitmap (ignore errors)
        DoEvents
        res = AVIStreamWrite(psCompressed, i, 1, bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
        If res <> AVIERR_OK Then GoTo error
        If docancel Then Exit Do
        sFile = Dir$()
    Loop

error:
'   Now close the file
    If (ps <> 0) Then Call AVIStreamClose(ps)
    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)
    If (pfile <> 0) Then Call AVIFileClose(pfile)
    Call AVIFileExit
    If (res <> AVIERR_OK) Then
        If res = AVIERR_BADFORMAT Or res = AVIERR_INTERNAL Then
            MsgBox "There was an error creating the AVI File." + vbCrLf + "Probably the choosen Video Compression does not support the File format or the input File is corrupt", vbInformation, App.Title
            frmAVI.sBar.SimpleText = "There was an error writing the file."
        Else
            MsgBox "There was an error creating the AVI File.", vbInformation, App.Title
            frmAVI.sBar.SimpleText = "There was an error writing the file."
        End If
    Else
        frmAVI.sBar.SimpleText = szOutputAVIFile + " created!"
    End If
    docancel = True
    
    
    cmdGo.Visible = True
    cmdCancel.Visible = False
    cmdClose.Enabled = True
    cmdSrc.Enabled = True
    cmdDest.Enabled = True
End Sub

Private Sub cmdSrc_Click()
    Dim sFile As String
    fSrc = SrcFile
    If fSrc Then
        Call loadPic(szInputFile, chkShowImage = vbChecked, True, False)
        fPath = GetFilePath(szInputFile)
        sFile = Dir$(fPath + "\*.*")
        telDir = 0
        Do While sFile <> ""
            telDir = telDir + 1
            sFile = Dir$()
        Loop
        cBar.SetParamFast 1, telDir, Left2Right, True, ShowCaption
    End If
    If fSrc And fDest Then cmdGo.Enabled = True
End Sub

Private Sub Form_Load()
    Set cBar = New ClsProgressBar2
    cBar.SetPictureBox = Picture1
    fSrc = False
    fDest = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set bmp = Nothing
    Set cBar = Nothing
End Sub

Function DestFile() As Boolean
    With CommonDialog1
        .DialogTitle = "Save AVI File"
        .CancelError = False
        .Filter = "AVI Files (*.avi)|*.avi"
        .DefaultExt = "avi"
        .filename = ""
        .ShowSave
    End With

    If Len(CommonDialog1.filename) = 0 Then
        DestFile = False
        Exit Function
    Else
        szOutputAVIFile = CommonDialog1.filename
        lblFileDest = GetFileBaseName(szOutputAVIFile) + "." + GetFileExtension(szOutputAVIFile)
        DestFile = True
    End If
End Function

Function SrcFile() As Boolean
    With CommonDialog1
        .DialogTitle = "Choose First Image File in directory"
        .CancelError = False
        .Filter = "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
        .DefaultExt = ""
        .filename = ""
        .ShowOpen
    End With

    If Len(CommonDialog1.filename) = 0 Then
        SrcFile = False
        Exit Function
    Else
        szInputFile = CommonDialog1.filename
        lblFileSrc = GetFileBaseName(szInputFile) + "." + GetFileExtension(szInputFile)
        SrcFile = True
    End If
End Function

Function loadPic(sFile As String, showIm As Boolean, clearImFirst As Boolean, saveBMP As Boolean) As String
    Dim xp As Single, yp As Single, xw As Single, yh As Single, xyi As Single, xys As Single
    Dim bmpFile As String
    imL.Picture = LoadPicture(sFile)
    If showIm Then
        xys = imShow.ScaleWidth / imShow.ScaleHeight
        xyi = imL.ScaleWidth / imL.ScaleHeight
        If xyi < xys Then
            xw = xyi * imShow.Width
            xp = (imShow.Width - xw) / 2
            yp = 0
            yh = imShow.Height
        Else
            xw = imShow.Width
            xp = 0
            yh = imShow.Height / xyi
            yp = (imShow.Height - yh) / 2
        End If
        If clearImFirst Then
            imShow.Picture = LoadPicture("")
            imShow.Refresh
        End If
        imShow.PaintPicture imL.Picture, xp, yp, xw, yh
        imShow.Refresh
    End If
    If saveBMP Then
        If GetFileExtension(sFile) <> "bmp" Then
            bmpFile = "c:\temp.bmp"
            SavePicture imL.Picture, bmpFile
        Else
            bmpFile = sFile
        End If
        loadPic = bmpFile
    Else
        loadPic = ""
    End If
End Function
