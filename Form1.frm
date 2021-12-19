VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Volume/CD Control"
   ClientHeight    =   3060
   ClientLeft      =   6615
   ClientTop       =   3765
   ClientWidth     =   2550
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   2550
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Volume Controls"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      Begin VB.VScrollBar vsVolume 
         Height          =   1455
         Left            =   480
         TabIndex        =   0
         Top             =   600
         Width           =   255
      End
      Begin VB.VScrollBar vsMic 
         Height          =   1455
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Microphone"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Volume"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   675
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblClose 
      BackColor       =   &H00000000&
      Caption         =   "Close Tray"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Form1.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblOpen 
      BackColor       =   &H00000000&
      Caption         =   "Open Tray"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Form1.frx":0316
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblQuit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      MouseIcon       =   "Form1.frx":0620
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblStop 
      BackColor       =   &H00000000&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      MouseIcon       =   "Form1.frx":092A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblPlay 
      BackColor       =   &H00000000&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      MouseIcon       =   "Form1.frx":0C34
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hmixer  As Long
Dim volCtrl As MIXERCONTROL ' Waveout volume control.
Dim micCtrl As MIXERCONTROL ' Microphone volume control.


Private Sub pShowError(lError As Long)
Dim lLen   As Long
Dim sError As String

sError = Space$(255)
lLen = Len(sError)

Call mciGetErrorString(lError, sError, lLen)
MsgBox Trim$(sError), vbCritical, "Error"

Call mciSendCommand(lDeviceID, MCI_CLOSE, 0, Null)
End Sub

Private Function fSetVolumeControl(ByVal hmixer As Long, _
    mxc As MIXERCONTROL, ByVal volume As Long) As Boolean
'
' This function sets the value for a volume control.
'
Dim rc   As Long
Dim mxcd As MIXERCONTROLDETAILS
Dim vol  As MIXERCONTROLDETAILS_UNSIGNED

With mxcd
    .item = 0
    .dwControlID = mxc.dwControlID
    .cbStruct = Len(mxcd)
    .cbDetails = Len(vol)
End With
'
' Allocate a buffer for the control value buffer.
'
hmem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hmem)
mxcd.cChannels = 1
vol.dwValue = volume
'
' Copy the data into the control value buffer.
'
Call CopyPtrFromStruct(mxcd.paDetails, vol, Len(vol))
'
' Set the control value.
'
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
Call GlobalFree(hmem)

If MMSYSERR_NOERROR = rc Then
    fSetVolumeControl = True
Else
    fSetVolumeControl = False
End If
End Function




Private Sub Form_Load()
'
'-----------------------------------------------
' Wave file related.
'-----------------------------------------------
'
Dim rc  As Long
Dim bOK As Boolean
' Open the mixer with deviceID 0.
'
rc = mixerOpen(hmixer, 0, 0, 0, 0)
If MMSYSERR_NOERROR <> rc Then
    MsgBox "Could not open the mixer.", vbCritical, "Volume Control"
    Exit Sub
End If
'
' Get the waveout volume control.
'
bOK = fGetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)
'
' If the function successfully gets the volume control,
' the maximum and minimum values are specified by
' lMaximum and lMinimum. Use them to set the scrollbar.
'
If bOK Then
    With vsVolume
        .Max = volCtrl.lMinimum
        .Min = volCtrl.lMaximum \ 2
        .SmallChange = 1000
        .LargeChange = 1000
    End With
End If
'
' Get the microphone volume control.
'
bOK = fGetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, micCtrl)

If bOK Then
    With vsMic
        .Max = micCtrl.lMinimum
        .Min = micCtrl.lMaximum \ 2
        .SmallChange = 1000
        .LargeChange = 1000
        .Enabled = True
    End With
End If
'
'-----------------------------------------------
' CD tray related.
'-----------------------------------------------
'
Dim lResult As Long

lFlags = MCI_OPEN_TYPE Or MCI_OPEN_SHAREABLE

tmciOpen.wDeviceID = 0
tmciOpen.lpstrDeviceType = "cdaudio"
'dss
Call mciSendCommand(lDeviceID, MCI_CLOSE, 0, Null)
lResult = mciSendCommand(0, MCI_OPEN, lFlags, tmciOpen)

If lResult <> 0 Then
    Call pShowError(lResult)
'    lblOpen.Enabled = False
'    lblClose.Enabled = False
Else
    lDeviceID = tmciOpen.wDeviceID
'    lblOpen.Enabled = True
'    lblClose.Enabled = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call mciSendCommand(lDeviceID, MCI_CLOSE, 0, Null)
Call lblStop_Click
Set Form1 = Nothing
End Sub


Private Sub lblClose_Click()
Dim lResult As Long

lResult = mciSendCommand(lDeviceID, MCI_SET, MCI_SET_DOOR_CLOSED, tmciSet)
If lResult <> 0 Then Call pShowError(lResult)
End Sub

Private Sub lblOpen_Click()
Dim lResult As Long

lResult = mciSendCommand(lDeviceID, MCI_SET, MCI_SET_DOOR_OPEN, tmciSet)
If lResult <> 0 Then Call pShowError(lResult)
End Sub


Private Sub LblQuit_Click()
Unload Me
End Sub

Private Sub lblPlay_Click()
Dim l          As Long
Dim lFlags     As Long
Dim sSoundName As String
'
'Open a wavefile and initialize the form.
'
On Error GoTo lblPlayError
With CommonDialog1
    .FileName = "*.wav"
    .DefaultExt = "wav"
    .Filter = "Wav (*.wav)"
    .FilterIndex = 1
    .Flags = cdlOFNPathMustExist Or cdlOFNFileMustExist
    .DialogTitle = "Select a Wave File"
    .CancelError = True
    .ShowOpen
    sSoundName = .FileName
End With

lFlags = SND_ASYNC Or SND_NODEFAULT Or SND_FILENAME
l = PlaySound(sSoundName, 0, lFlags)

lblPlayError:
End Sub

Private Sub lblStop_Click()
Dim l          As Long
Dim lFlags     As Long

lFlags = SND_ASYNC Or SND_NODEFAULT Or SND_FILENAME
l = PlaySound("", 0, lFlags)
End Sub

Private Sub vsMic_Change()
Dim lVol As Long

lVol = CLng(vsMic.Value) * 2
Call fSetVolumeControl(hmixer, micCtrl, lVol)
End Sub


Private Sub vsMic_Scroll()
Dim lVol As Long

lVol = CLng(vsMic.Value) * 2
Call fSetVolumeControl(hmixer, micCtrl, lVol)
End Sub


Private Sub vsVolume_Change()
Dim lVol As Long

lVol = CLng(vsVolume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub


Private Sub vsVolume_Scroll()
Dim lVol As Long

lVol = CLng(vsVolume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub


