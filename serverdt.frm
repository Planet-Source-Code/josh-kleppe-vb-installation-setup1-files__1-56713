VERSION 5.00
Begin VB.Form frmRemoteServerDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   4545
   ClientLeft      =   3195
   ClientTop       =   2400
   ClientWidth     =   7800
   ControlBox      =   0   'False
   Icon            =   "serverdt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7800
   Begin VB.CommandButton cmdCancel 
      Caption         =   "#"
      Height          =   375
      Left            =   5580
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   3930
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "#"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3540
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   3930
      Width           =   1935
   End
   Begin VB.ComboBox cboNetworkProtocol 
      Height          =   300
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3165
      Width           =   5100
   End
   Begin VB.TextBox txtNetworkAddress 
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   2535
      Width           =   5100
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   225
      TabIndex        =   7
      Top             =   1395
      Width           =   7290
      Begin VB.Label lblServerName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   240
         Width           =   7020
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblNetworkProtocol 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   3165
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNetworkAddress 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   2535
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRemoteServerDetails 
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   7020
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRemoteServerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private m_fNetworkAddressSpecified As Boolean
Private m_fNetworkProtocolSpecified As Boolean
Private m_fDCOM As Boolean

Private Declare Function RpcNetworkIsProtseqValid Lib "rpcrt4.dll" Alias "RpcNetworkIsProtseqValidA" (ByVal strProtseq As String) As Long

' Determines whether a given protocol sequence is supported and available on this machine
Function fIsProtocolSeqSupported(ByVal strProto As String, ByVal strProtoFriendlyName) As Boolean
    Const RPC_S_OK = 0&
    Const RPC_S_PROTSEQ_NOT_SUPPORTED = 1703&
    Const RPC_S_INVALID_RPC_PROTSEQ = 1704&

    Dim rcps As Long
    Static fUnexpectedErr As Boolean

    On Error Resume Next

    fIsProtocolSeqSupported = False
    rcps = RpcNetworkIsProtseqValid(strProto)

    Select Case rcps
        Case RPC_S_OK
            fIsProtocolSeqSupported = True
        Case RPC_S_PROTSEQ_NOT_SUPPORTED
            LogNote ResolveResString(resNOTEPROTOSEQNOTSUPPORTED, "|1", strProto, "|2", strProtoFriendlyName)
        Case RPC_S_INVALID_RPC_PROTSEQ
            LogWarning ResolveResString(resNOTEPROTOSEQINVALID, "|1", strProto, "|2", strProtoFriendlyName)
        Case Else
            If Not fUnexpectedErr Then
                MsgWarning ResolveResString(resPROTOSEQUNEXPECTEDERR), vbOKOnly Or vbInformation, gstrTitle
                If gfNoUserInput Then
                    '
                    ' This is probably redundant since this form should never
                    ' be shown if we are running in silent or SMS mode.
                    '
                    ExitSetup frmRemoteServerDetails, gintRET_FATAL
                End If
                fUnexpectedErr = True
            End If
        'End Case
    End Select
End Function

Private Sub cboNetworkProtocol_Click()
    cmdOK.Enabled = fValid()
End Sub

Private Sub cmdCancel_Click()
    ExitSetup frmRemoteServerDetails, gintRET_EXIT
End Sub

Private Sub cmdOK_Click()
    Hide
End Sub

Private Sub Form_Load()
    Dim fMoveControlsUp As Boolean 'Whether or not to move controls up to fill in an empty space
    Dim yTopCutoff As Integer 'We will move all controls lower down than this y value
    
    SetFormFont Me
    Caption = ResolveResString(resREMOTESERVERDETAILSTITLE)
    lblRemoteServerDetails.Caption = ResolveResString(resREMOTESERVERDETAILSLBL)
    lblNetworkAddress.Caption = ResolveResString(resNETWORKADDRESS)
    lblNetworkProtocol.Caption = ResolveResString(resNETWORKPROTOCOL)
    cmdOK.Caption = ResolveResString(resOK)
    cmdCancel.Caption = ResolveResString(resCANCEL)
    '
    ' We don't care about protocols if this is DCOM.
    '
    If Not m_fDCOM Then
        FillInProtocols
    End If
    
    'Now we selectively turn on/off the available controls depending on how
    '  much information we need from the user.
    If m_fNetworkAddressSpecified Then
        'The network address has already been filled in, so we can hide this
        '  control and move all the other controls up
        txtNetworkAddress.Visible = False
        lblNetworkAddress.Visible = False
        fMoveControlsUp = True
        yTopCutoff = txtNetworkAddress.Top
    ElseIf m_fNetworkProtocolSpecified Or m_fDCOM Then
        'The network protocol has already been filled in, so we can hide this
        '  control and move all the other controls up
        cboNetworkProtocol.Visible = False
        lblNetworkProtocol.Visible = False
        fMoveControlsUp = True
        yTopCutoff = cboNetworkProtocol.Top
    End If
    
    If fMoveControlsUp Then
        'Find out how much to move the controls up
        Dim yDiff As Integer
        yDiff = cboNetworkProtocol.Top - txtNetworkAddress.Top
        
        Dim c As Control
        For Each c In Controls
            If c.Top > yTopCutoff Then
                c.Top = c.Top - yDiff
            End If
        Next c
        
        'Finally, shrink the form
        Height = Height - yDiff
    End If
    
    'Center the form
    Top = (Screen.Height - Height) \ 2
    Left = (Screen.Width - Width) \ 2
End Sub

'-----------------------------------------------------------
' SUB: GetServerDetails
'
' Requests any missing information about a remote server from
' the user.
'
' Input:
'   [strRegFile] - the name of the remote registration file
'   [strNetworkAddress] - the network address, if known
'   [strNetworkProtocol] - the network protocol, if known
'   [fDCOM] - if true, this component is being accessed via
'             distributed com and not Remote automation.  In
'             this case, we don't need the network protocol or
'             Authentication level.
'
' Ouput:
'   [strNetworkAddress] - the network address either passed
'                         in or obtained from the user
'   [strNetworkProtocol] - the network protocol either passed
'                          in or obtained from the user
'-----------------------------------------------------------
'
Public Sub GetServerDetails( _
    ByVal strRegFile As String, _
    strNetworkAddress As String, _
    strNetworkProtocol As String, _
    fDCOM As Boolean _
)
    Dim i As Integer
    Dim strServerName As String
    
    'See if anything is missing
    m_fNetworkAddressSpecified = (strNetworkAddress <> "")
    m_fNetworkProtocolSpecified = (strNetworkProtocol <> "")
    m_fDCOM = fDCOM
    
    If m_fNetworkAddressSpecified And (m_fNetworkProtocolSpecified Or m_fDCOM) Then
        'Both the network address and protocol sequence have already
        'been specified in SETUP.LST.  There is no need to ask the
        'user for more information.
        
        'However, we do need to check that the protocol sequence specified
        'in SETUP.LST is actually installed and available on this machine
        '(Remote Automation only).
        '
        If Not m_fDCOM Then
            CheckSpecifiedProtocolSequence strNetworkProtocol, strGetServerName(strRegFile)
        End If
        
        Exit Sub
    End If
    
    strServerName = strGetServerName(strRegFile)
''''''''''''''''''''''''''JPK 1/12/98 '''''''''''''''''''''''''''
    If gNonSilentInstall Then Load Me
    If gNonSilentInstall Then lblServerName.Caption = strServerName
''''''''''''''''''''''''''JPK 1/12/98 '''''''''''''''''''''''''''
    
    If Not gfNoUserInput Then
        '
        ' Show the form and extract necessary information from the user
        '
 ''''''''''''''''''''''''''JPK 1/12/98 '''''''''''''''''''''''''''
    If gNonSilentInstall Then Show vbModal
''''''''''''''''''''''''''JPK 1/12/98 '''''''''''''''''''''''''''
    Else
        '
        ' Since this is silent, simply accept the first one on
        ' the list.
        '
        ' Note that we know there is at least 1 protocol in the
        ' list or else the program would have aborted in
        ' the Form_Load code when it called FillInProtocols().
        '
        cboNetworkProtocol.ListIndex = 0
    End If
    
    If m_fNetworkProtocolSpecified And Not m_fDCOM Then
        'The network protocol sequence had already been specified
        'in SETUP.LST.  We need to check that the protocol sequence specified
        'in SETUP.LST is actually installed and available on this machine
        '(32-bit only).
        CheckSpecifiedProtocolSequence strNetworkProtocol, strGetServerName(strRegFile)
    End If
    
    If Not m_fNetworkAddressSpecified Then
        strNetworkAddress = txtNetworkAddress
    End If
    If Not m_fNetworkProtocolSpecified And Not m_fDCOM Then
        strNetworkProtocol = gProtocol(cboNetworkProtocol.ListIndex + 1).strName
    End If
    
''''''''''''''''''''''''''JPK 1/12/98 '''''''''''''''''''''''''''
    If gNonSilentInstall Then Unload Me
''''''''''''''''''''''''''JPK 1/12/98 '''''''''''''''''''''''''''
End Sub

'-----------------------------------------------------------
' SUB: FillInProtocols
'
' Fills in the protocol combo with the available protocols from
'   setup.lst
'-----------------------------------------------------------
Private Sub FillInProtocols()
    Dim i As Integer
    Dim fSuccessReading As Boolean
    
    cboNetworkProtocol.Clear
    fSuccessReading = ReadProtocols(gstrSetupInfoFile, gstrINI_SETUP)
    If Not fSuccessReading Or gcProtocols <= 0 Then
        MsgError ResolveResString(resNOPROTOCOLSINSETUPLST), vbExclamation Or vbOKOnly, gstrTitle
        ExitSetup frmRemoteServerDetails, gintRET_FATAL
    End If
    
    For i = 1 To gcProtocols
        If fIsProtocolSeqSupported(gProtocol(i).strName, gProtocol(i).strFriendlyName) Then
            cboNetworkProtocol.AddItem gProtocol(i).strFriendlyName
        End If
    Next i

    If cboNetworkProtocol.ListCount > 0 Then
        'We were successful in finding at least one protocol available on this machine
        Exit Sub
    End If
    
    'None of the protocols specified in SETUP.LST are available on this machine.  We need
    'to let the user know what's wrong, including which protocol(s) were expected.
    MsgError ResolveResString(resNOPROTOCOLSSUPPORTED1), vbExclamation Or vbOKOnly, gstrTitle
    '
    ' Don't log the rest if this is SMS.  Ok for silent mode since
    ' silent can take more than 255 characters.
    '
    If Not gfSMS Then
        Dim strMsg As String
        strMsg = ResolveResString(resNOPROTOCOLSSUPPORTED2) & LF$
    
        For i = 1 To gcProtocols
            strMsg = strMsg & LF$ & Chr$(9) & gProtocol(i).strFriendlyName
        Next i
        
        MsgError strMsg, vbExclamation Or vbOKOnly, gstrTitle
    End If
    ExitSetup frmRemoteServerDetails, gintRET_FATAL
End Sub

'-----------------------------------------------------------
' SUB: strGetServerName
'
' Given a remote server registration file, retrieves the
'   friendly name of the server
'-----------------------------------------------------------
Private Function strGetServerName(ByVal strRegFilename As String) As String
    Const strKey = "AppDescription="
    Dim strLine As String
    Dim iFile As Integer
    
    On Error GoTo DoErr
    
    'This will have to do if we can't find the friendly name
    strGetServerName = GetFileName(strRegFilename)
    
    iFile = FreeFile
    Open strRegFilename For Input Access Read Lock Read Write As #iFile
    While Not EOF(iFile)
        Line Input #iFile, strLine
        If Left$(strLine, Len(strKey)) = strKey Then
            'We've found the line with the friendly server name
            Dim strName As String
            strName = Mid$(strLine, Len(strKey) + 1)
            If strName <> "" Then
                strGetServerName = strName
            End If
            Close iFile
            Exit Function
        End If
    Wend
    Close iFile
    
    Exit Function
DoErr:
    strGetServerName = ""
End Function

Private Sub txtNetworkAddress_Change()
    cmdOK.Enabled = fValid()
End Sub

'Returns True iff the inputs are valid
Private Function fValid() As Boolean
    fValid = True
    '
    ' If this is dcom, we don't care about the network protocol.
    '
    If m_fDCOM = False Then
        If Not m_fNetworkProtocolSpecified And (cboNetworkProtocol.ListIndex < 0) Then
            fValid = False
        End If
    End If
    
    If Not m_fNetworkAddressSpecified And (txtNetworkAddress = "") Then
        fValid = False
    End If
End Function

Private Sub CheckSpecifiedProtocolSequence(ByVal strNetworkProtocol As String, ByVal strFriendlyServerName As String)
        'Attempt to find the friendly name of this protocol from the list in SETUP.LST
        Dim fSuccessReading As Boolean
        Dim strFriendlyName As String
        Dim i As Integer
        
        strFriendlyName = strNetworkProtocol 'This will have to do if we can't find anything better
        
        fSuccessReading = ReadProtocols(gstrSetupInfoFile, gstrINI_SETUP)
        If fSuccessReading And gcProtocols > 0 Then
            For i = 1 To gcProtocols
                If gProtocol(i).strName = strNetworkProtocol Then
                    strFriendlyName = gProtocol(i).strFriendlyName
                    Exit For
                End If
            Next i
        End If
        
        'Now check to see if this protocol is available
        If fIsProtocolSeqSupported(strNetworkProtocol, strFriendlyName) Then
            'OK
            Exit Sub
        Else
            'Nope, not supported.  Give an informational message about what to do, then continue with setup.
Retry:
            If gfNoUserInput Or MsgError( _
                ResolveResString(resSELECTEDPROTONOTSUPPORTED, "|1", strFriendlyServerName, "|2", strFriendlyName), _
                vbInformation Or vbOKCancel, _
                gstrTitle) _
              = vbCancel Then
                '
                ' The user chose cancel.  Give them a chance to exit (if this isn't a silent or sms install;
                ' otherwise any call to ExitSetup is deemed fatal.
                '
                ExitSetup frmRemoteServerDetails, gintRET_EXIT
                GoTo Retry
            Else
                'The user chose OK.  Continue with setup.
                Exit Sub
            End If
        End If
End Sub

