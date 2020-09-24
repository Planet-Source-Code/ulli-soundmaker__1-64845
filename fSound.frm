VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fSound 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Soundmaker"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   ForeColor       =   &H8000000F&
   Icon            =   "fSound.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4710
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picShow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00D0FF00&
      Height          =   1560
      Left            =   75
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   25
      Top             =   90
      Width           =   4560
   End
   Begin VB.CheckBox ckDecay 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00C0FFFF&
      Caption         =   "Decay"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   3030
      Value           =   1  'Aktiviert
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog cdlSave 
      Left            =   4200
      Top             =   3765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save as"
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   450
      Left            =   2850
      TabIndex        =   21
      Top             =   3780
      Width           =   975
   End
   Begin VB.CheckBox chkLink 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Link"
      Height          =   240
      Left            =   3810
      TabIndex        =   12
      Top             =   3360
      Value           =   1  'Aktiviert
      Width           =   600
   End
   Begin VB.HScrollBar scrVolume 
      Height          =   120
      Index           =   1
      LargeChange     =   3277
      Left            =   998
      SmallChange     =   328
      TabIndex        =   11
      Top             =   3495
      Value           =   16384
      Width           =   2715
   End
   Begin VB.HScrollBar scrVolume 
      Height          =   120
      Index           =   0
      LargeChange     =   3277
      Left            =   998
      SmallChange     =   328
      TabIndex        =   9
      Top             =   3345
      Value           =   16384
      Width           =   2715
   End
   Begin VB.OptionButton optWF 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Squarewave"
      Height          =   195
      Index           =   5
      Left            =   2430
      TabIndex        =   18
      Top             =   3015
      Width           =   1215
   End
   Begin VB.OptionButton optWF 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sawtooth"
      Height          =   195
      Index           =   4
      Left            =   2430
      TabIndex        =   17
      Top             =   2805
      Width           =   975
   End
   Begin VB.OptionButton optWF 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sine with harmonics"
      Height          =   195
      Index           =   3
      Left            =   2430
      TabIndex        =   16
      Top             =   2595
      Width           =   1710
   End
   Begin VB.OptionButton optWF 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pure Sine w/ White Noise"
      Height          =   195
      Index           =   2
      Left            =   2430
      TabIndex        =   15
      Top             =   2370
      Width           =   2160
   End
   Begin VB.OptionButton optWF 
      BackColor       =   &H00C0FFFF&
      Caption         =   "White Noise"
      Height          =   195
      Index           =   1
      Left            =   2430
      TabIndex        =   14
      Top             =   2145
      Width           =   1185
   End
   Begin VB.OptionButton optWF 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pure Sine"
      Height          =   195
      Index           =   0
      Left            =   2430
      TabIndex        =   13
      Top             =   1920
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Play Again"
      Enabled         =   0   'False
      Height          =   450
      Left            =   1868
      TabIndex        =   20
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton cmdGen 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Generate"
      Height          =   450
      Left            =   900
      TabIndex        =   19
      Top             =   3780
      Width           =   975
   End
   Begin VB.TextBox txtF1 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   1035
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "440"
      Top             =   2280
      Width           =   720
   End
   Begin VB.TextBox txtF2 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   1035
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "440"
      Top             =   2640
      Width           =   720
   End
   Begin VB.TextBox txtDuration 
      Alignment       =   1  'Rechts
      Height          =   300
      Left            =   1035
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "1000"
      Top             =   1935
      Width           =   720
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   " Four cycles or full wave depending on which is less "
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   150
      Index           =   9
      Left            =   945
      TabIndex        =   26
      Top             =   1665
      Width           =   2805
   End
   Begin VB.Line ln 
      BorderColor     =   &H0000C0C0&
      Index           =   2
      X1              =   105
      X2              =   330
      Y1              =   1740
      Y2              =   1800
   End
   Begin VB.Line ln 
      BorderColor     =   &H0000C0C0&
      Index           =   1
      X1              =   105
      X2              =   330
      Y1              =   1740
      Y2              =   1680
   End
   Begin VB.Line ln 
      BorderColor     =   &H0000C0C0&
      Index           =   3
      X1              =   4560
      X2              =   4335
      Y1              =   1740
      Y2              =   1680
   End
   Begin VB.Line ln 
      BorderColor     =   &H0000C0C0&
      Index           =   4
      X1              =   4575
      X2              =   4335
      Y1              =   1740
      Y2              =   1800
   End
   Begin VB.Line ln 
      BorderColor     =   &H0000C0C0&
      Index           =   0
      X1              =   90
      X2              =   4590
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "right"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   8
      Left            =   660
      TabIndex        =   10
      Top             =   3465
      Width           =   270
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "left"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   7
      Left            =   690
      TabIndex        =   8
      Top             =   3315
      Width           =   195
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz"
      Height          =   195
      Index           =   6
      Left            =   1845
      TabIndex        =   24
      Top             =   2685
      Width           =   195
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz"
      Height          =   195
      Index           =   5
      Left            =   1845
      TabIndex        =   23
      Top             =   2325
      Width           =   195
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "mSec"
      Height          =   195
      Index           =   4
      Left            =   1845
      TabIndex        =   22
      Top             =   1980
      Width           =   405
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      Height          =   195
      Index           =   3
      Left            =   105
      TabIndex        =   7
      Top             =   3360
      Width           =   525
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Freq Right"
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   4
      Top             =   2700
      Width           =   735
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Freq Left"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   2325
      Width           =   630
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Duration"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   1980
      Width           =   600
   End
End
Attribute VB_Name = "fSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" () ':) Line inserted by Formatter
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_MEMORY    As Long = 4
Private Const SND_ASYNC     As Long = 1

Private Type tHeader
    RIFF        As Long         ' "RIFF"
    LenR        As Long         '  size of following segment
    WAVE        As Long         ' "WAVE"
    fmt         As Long         ' "fmt
    LenChu      As Long         '  chunksize
    AudioFmt    As Integer      '  wav format = 1
    NumChan     As Integer      '  nunber of channels
    SpS         As Long         '  samples per second
    BpS         As Long         '  bytes per second
    Blk         As Integer      '  block align (bytes per sample)
    LenSam      As Integer      '  bits per sample
    data        As Long         ' "data"
    LenData     As Long         '  length of datastream
End Type
Private Header                  As tHeader

Private Type tSample
    LeftRight(0 To 1)           As Integer
End Type
Private Samples()               As tSample

Private Const SamplesPerSecond  As Long = 44100     '22050 cuts filesize in half with loss of fidelity
Private Const BitsPerSample     As Long = 16        'do not alter
Private Const NumberOfChannels  As Long = 2         'stereo - do not alter
Private WaveForm                As Long

Private SoundFile()             As Byte

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Function CastToLong(Text As String) As Long

  Dim i As Long

    For i = 4 To 1 Step -1
        CastToLong = CastToLong * 256 + Asc(Mid$(Text, i, 1))
    Next i

End Function

Private Sub chkLink_Click()

  Dim Ena As Long

    Ena = cmdPlay.Enabled
    scrVolume_Scroll 0
    cmdPlay.Enabled = Ena
    cmdSave.Enabled = Ena

End Sub

Private Sub ckDecay_Click()

    cmdPlay.Enabled = False
    cmdSave.Enabled = False

End Sub

Private Sub cmdGen_Click()

    Enabled = False
    MakeSoundFile Val(txtF1.Text), scrVolume(0), Val(txtF2.Text), scrVolume(1), Val(txtDuration.Text) / 1000, WaveForm, ckDecay = vbChecked
    Enabled = True
    cmdSave.Enabled = True
    cmdPlay.Enabled = True

End Sub

Private Sub cmdPlay_Click()

    Enabled = False
    PlaySoundData SoundFile(1), 0, SND_MEMORY
    Enabled = True

End Sub

Private Sub cmdSave_Click()

  Dim Abandon   As Boolean
  Dim hFile     As Long

    With cdlSave
        .InitDir = App.Path
        .Filter = "Sound File Type 1 (*.wav)|*.wav"
        .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt Or cdlOFNPathMustExist
        On Error Resume Next
            .ShowSave
            Abandon = (Err = cdlCancel)
        On Error GoTo 0
        If Not Abandon Then
            hFile = FreeFile
            Open .FileName For Binary As hFile
            Put hFile, , SoundFile
            Close hFile
        End If
    End With 'CDLSAVE

End Sub

Private Sub MakeSoundFile(ByVal FreqLeft As Long, ByVal VolLeft As Double, ByVal FreqRight As Long, ByVal VolRight As Double, ByVal Duration As Single, ByVal WaveType As Long, ByVal Decay As Boolean)

  Dim Side              As Long
  Dim Ptr               As Long
  Dim NumSamples        As Long
  Dim SamplesPerCycle   As Double
  Dim Time              As Double
  Dim DeltaTime         As Double
  Dim Vol               As Double
  Dim DeltaVol          As Double
  Dim Frq               As Long
  Dim Omega             As Double
  Dim Sample            As Integer
  Dim dW                As Long     'draw width
  Dim dH                As Long     'draw height
  Dim sH                As Double   'scale horizontal
  Dim sV                As Double   'scale vertival
  Dim dZ(0 To 1)        As Long     'draw zeroline left/right

    With picShow
        .Cls
        dW = .ScaleWidth
        dH = .ScaleHeight / 2 - 1
        dZ(0) = .ScaleHeight / 4
    End With 'PICSHOW
    dZ(1) = dZ(0) + dH
    sV = dH / 70000
    DoEvents
    Duration = Abs(Duration)
    NumSamples = SamplesPerSecond * Duration
    If NumSamples = 0 Then
        NumSamples = 1
    End If
    DeltaTime = 1 / SamplesPerSecond
    WaveType = WaveType Mod 6

    ReDim Samples(0 To NumSamples - 1)
    For Side = 0 To 1  '0 - left  1 - right
        picShow.Line (dW, dZ(Side))-(-1, dZ(Side)), vbRed 'zero line and initial draw point
        If Side = 0 Then
            Frq = Abs(FreqLeft) Mod (SamplesPerSecond / 2 + 1)
            Vol = VolLeft And &H7FFF
          Else 'NOT SIDE...
            Frq = Abs(FreqRight) Mod (SamplesPerSecond / 2 + 1)
            Vol = VolRight And &H7FFF
        End If
        Omega = 8 * Atn(1) * Frq
        If Decay Then
            DeltaVol = Vol / NumSamples
        End If
        If Frq = 0 Then
            SamplesPerCycle = NumSamples
          Else 'NOT FRQ...
            SamplesPerCycle = SamplesPerSecond / Frq
        End If
        sH = SamplesPerCycle * 4
        If sH > NumSamples Then
            sH = NumSamples
        End If

        Time = -1.13378684807256E-05 'little phase shift becomes 90° at 22050 Hz
        For Ptr = 0 To NumSamples - 1
            Select Case WaveType
              Case 0 'PureSine
                Sample = Vol * Sin(Omega * Time)
              Case 1 ' WhiteNoise
                Sample = Vol * (Rnd - Rnd)
              Case 2 'PureSine with white noise
                Sample = Vol * Sin(Omega * Time) * (1 - Rnd / 4)
              Case 3 'Sine with Harmonics
                Sample = Vol * Sin(Omega * Time) ^ 3 'exponent must be odd to preserve the sign of the sample
              Case 4 'Sawtooth
                Sample = Vol / SamplesPerCycle * 2 * (Ptr Mod SamplesPerCycle) - Vol
              Case 5 'Squarewave
                Sample = Vol * Sgn(Sin(Omega * Time))
            End Select
            Samples(Ptr).LeftRight(Side) = Sample
            If Ptr <= sH + 2 Then
                picShow.Line -(dW * Ptr / sH - 1, dZ(Side) - Sample * sV)
            End If
            Time = Time + DeltaTime
            Vol = Vol - DeltaVol
    Next Ptr, Side

    With Header
        .RIFF = CastToLong("RIFF")
        .WAVE = CastToLong("WAVE")
        .fmt = CastToLong("fmt ")
        .data = CastToLong("data")
        .NumChan = NumberOfChannels
        .SpS = SamplesPerSecond
        .LenSam = BitsPerSample
        .LenChu = 16
        .AudioFmt = 1
        .Blk = .NumChan * .LenSam / 8
        .BpS = .SpS * .Blk
        .LenData = NumSamples * .NumChan * .LenSam / 8
        .LenR = Len(Header) + .LenData - 8
        ReDim SoundFile(1 To Len(Header) + .LenData)
        MemCopy VarPtr(SoundFile(1)), VarPtr(Header), Len(Header)
        MemCopy VarPtr(SoundFile(1)) + Len(Header), VarPtr(Samples(0)), .LenData
    End With 'HEADER

    Erase Samples
    DoEvents
    cmdPlay_Click

End Sub

Private Sub optWF_Click(Index As Integer)

    WaveForm = Index
    cmdPlay.Enabled = False
    cmdSave.Enabled = False
    txtF1.Enabled = (Index <> 1)
    txtF2.Enabled = (Index <> 1)

End Sub

Private Sub scrVolume_Change(Index As Integer)

    If chkLink = vbChecked Then
        scrVolume(1 - Index) = scrVolume(Index)
    End If
    cmdPlay.Enabled = False
    cmdSave.Enabled = False

End Sub

Private Sub scrVolume_Scroll(Index As Integer)

    scrVolume_Change Index

End Sub

Private Sub txtDuration_Change()

    txtF1_Change

End Sub

Private Sub txtF1_Change()

    cmdPlay.Enabled = False
    cmdSave.Enabled = False

End Sub

Private Sub txtF2_Change()

    txtF1_Change

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Mrz-30 10:57)  Decl: 36  Code: 232  Total: 268 Lines
':) CommentOnly: 0 (0%)  Commented: 37 (13,8%)  Empty: 55 (20,5%)  Max Logic Depth: 4
