VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Email with Option Buttons"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmEMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   7455
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   350
      Left            =   105
      TabIndex        =   15
      Top             =   6150
      Width           =   1305
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Left            =   105
      TabIndex        =   23
      Top             =   60
      Width           =   7260
      Begin VB.ComboBox cboAddressBooks 
         Height          =   315
         Left            =   1395
         TabIndex        =   1
         Top             =   450
         Width           =   2970
      End
      Begin VB.Label lblBooks 
         BackStyle       =   0  'Transparent
         Caption         =   "&Address Books:"
         Height          =   255
         Left            =   210
         TabIndex        =   0
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Selecting and address book entry will populate the Recipients drop-down, box."
         Height          =   300
         Left            =   105
         TabIndex        =   24
         Top             =   180
         Width           =   7065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   105
      TabIndex        =   19
      Top             =   990
      Width           =   7260
      Begin VB.CommandButton cmdLoadRecipientsFromFile 
         Caption         =   "Load &From File"
         Height          =   350
         Left            =   4455
         TabIndex        =   4
         Top             =   510
         Width           =   1305
      End
      Begin VB.CommandButton cmdSaveRecipientsToFile 
         Caption         =   "Save &To File"
         Height          =   350
         Left            =   5820
         TabIndex        =   5
         Top             =   510
         Width           =   1305
      End
      Begin VB.ComboBox cboContacts 
         Height          =   315
         Left            =   1395
         TabIndex        =   3
         Top             =   525
         Width           =   2970
      End
      Begin VB.Label Label2 
         Caption         =   "Enter a semi-colon delimited list of recipients, select from the drop-down, or load recipients from a file."
         Height          =   300
         Left            =   105
         TabIndex        =   22
         Top             =   180
         Width           =   7065
      End
      Begin VB.Label lblTo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Recipents:"
         Height          =   255
         Left            =   210
         TabIndex        =   2
         Top             =   540
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog cdlgMain 
      Left            =   90
      Top             =   6090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   350
      Left            =   6045
      TabIndex        =   17
      Top             =   6150
      Width           =   1305
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   350
      Left            =   4680
      TabIndex        =   16
      Top             =   6150
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Caption         =   "Email Contents"
      Height          =   4020
      Left            =   105
      TabIndex        =   20
      Top             =   2055
      Width           =   7260
      Begin VB.TextBox txtMessage 
         Height          =   1725
         Left            =   1395
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   630
         Width           =   5685
      End
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   1395
         TabIndex        =   7
         Top             =   255
         Width           =   5685
      End
      Begin VB.CheckBox chkUseVotingOptions 
         Caption         =   "&Use Voting Buttons:"
         Height          =   375
         Left            =   210
         TabIndex        =   10
         Top             =   2490
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.Frame fraVotingOptions 
         Caption         =   "Voting Options"
         Height          =   1425
         Left            =   1410
         TabIndex        =   21
         Top             =   2430
         Width           =   5700
         Begin VB.TextBox txtVotingOptions 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   525
            Width           =   4035
         End
         Begin VB.CommandButton cmdSaveOptionListToFile 
            Caption         =   "Save &To File"
            Height          =   350
            Left            =   4260
            TabIndex        =   14
            Top             =   930
            Width           =   1305
         End
         Begin VB.CommandButton cmdLoadOptionListFromFile 
            Caption         =   "Load &From File"
            Height          =   350
            Left            =   4260
            TabIndex        =   13
            Top             =   540
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "&Enter a semi-colon delimited list of options, or load options from a file."
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   255
            Width           =   5385
         End
      End
      Begin VB.Label lblBody 
         BackStyle       =   0  'Transparent
         Caption         =   "&Message body:"
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   690
         Width           =   1260
      End
      Begin VB.Label lblSubject 
         BackStyle       =   0  'Transparent
         Caption         =   "Su&bject:"
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   330
         Width           =   1260
      End
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -15
      TabIndex        =   18
      Top             =   6675
      Width           =   7470
   End
End
Attribute VB_Name = "frmEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   '__________________________________________________
   ' Module : FrmEMail
   ' Desc   : Allows sending email with option buttons.
   '__________________________________________________
   ' History
   '
   ' CDK: 20010105: Took an existing application written by
   '        Dylan Morley and posted on Planet Source.
   '        I modified it by allowing user to enter
   '        recipients without having to select from
   '        the outlook address book list.  user may
   '        load a recipient list from a file and save
   '        recipient to a file.
   '        Users may also now add voting buttons to
   '        their email.
   '        Added ability to clear form
   '        added hot keys and taborder
   '        changed interface look and feel
   '
   '        Thanks Dylan for the start - i needed this
   '            to help us at work figure out what we're
   '            doing for lunch every day.  Every day
   '            i send a list of options out to my
   '            lunch cronies, we go with the place
   '            with the highest vote.
   '__________________________________________________



Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260



Private Const mcsDefaultOptionList As String = "_DefaultOptionList.lst"
Private Const mcsDefaultRecipientList As String = "_DefaultRecipientList.lst"
Private Const mcsDefaultExtention As String = ".lst"
Private Const mcsDefaultDialogTitle  As String = "Voting Options"
Private Const mcsFilterForCommonDialog As String = "Voting Option List Files (*.lst)|*.lst"
Private Const mcsFilterForContentsFiles As String = "Little Helper Contents Files (*.lht)|*.lht"

Private Enum eCommonDialogConst
    ecdlShowOpen = 1
    ecdlShowSave = 2
    ecdlShowColor = 3
    ecdlShowFont = 4
    ecdlShowPrinter = 5
    ecdlShowWinHelp32 = 6
End Enum

Private Const mcsNoFileSelected As String = "No file was selected, do you want to use the default Contents?"


Private moOutlookApplication As Outlook.Application
Private moOutlookNamespace As Outlook.NameSpace
Private fErrorFlag As Boolean

Private Function InitializeOutlook() As Boolean

   '__________________________________________________
   ' Scope  :
   ' Type   : Function
   ' Name   : InitializeOutlook
   ' Params :
   ' Returns: Boolean
   ' Desc   : The Function uses parameters  for InitializeOutlook and returns Boolean.
   '__________________________________________________
   ' History
   ' CDK: 20010105: Added Error Trapping & Comments
   '__________________________________________________

   
    On Error GoTo Init_Err
    
    'Open an obj-var of type Outlook.Application
    Set moOutlookApplication = New Outlook.Application             ' Application object.
    
    'Use the Outlook.Application obj-var to create
    'a valid namespace
    Set moOutlookNamespace = moOutlookApplication.GetNamespace("MAPI")  ' Namespace object.
    
    InitializeOutlook = True
    
    Exit Function

Init_Err:
   InitializeOutlook = False
   
End Function


Private Sub chkUseVotingOptions_Click()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : chkUseVotingOptions_Click
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for chkUseVotingOptions_Click and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "chkUseVotingOptions_Click"

    Me.fraVotingOptions.Enabled = Me.chkUseVotingOptions.Value

Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

Private Sub cmdCancel_Click()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : cmdCancel_Click
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for cmdCancel_Click and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error Resume Next
    

'Unload and halt
    Unload Me

End Sub


Private Sub cmdLoadOptionListFromFile_Click()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : cmdLoadOptionListFromFile_Click
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for cmdLoadOptionListFromFile_Click and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "cmdLoadOptionListFromFile_Click"

    
    Dim sContents As String
    sContents = GetContentsFileName(ecdlShowOpen, App.Title & mcsDefaultOptionList, "Voting Options")
        
    'if no file selected, use default
    If sContents & "" = "" Then
        If MsgBox(mcsNoFileSelected, vbQuestion + vbYesNoCancel) = vbYes Then
            sContents = App.Path + "\" + App.Title + mcsDefaultOptionList
        End If
    End If
    
    Call LoadContentsFromFile(sContents, Me.txtVotingOptions)


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


Private Sub cmdLoadRecipientsFromFile_Click()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : cmdLoadRecipientsFromFile_Click
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for cmdLoadRecipientsFromFile_Click and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "cmdLoadRecipientsFromFile_Click"

    
    Dim sContents As String
    sContents = GetContentsFileName(ecdlShowOpen, App.Title & mcsDefaultRecipientList, "Recipient List")
        
    'if no file selected, use default
    If sContents & "" = "" Then
        If MsgBox(mcsNoFileSelected, vbQuestion + vbYesNoCancel) = vbYes Then
            sContents = App.Path + "\" + App.Title + mcsDefaultRecipientList
        End If
    End If
    
    Call LoadContentsFromFile(sContents, Me.cboContacts)


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub




Private Sub cmdSaveOptionListToFile_Click()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : cmdSaveOptionListToFile_Click
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for cmdSaveOptionListToFile_Click and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "cmdSaveOptionListToFile_Click"

    

    Dim sContents As String
    sContents = GetContentsFileName(ecdlShowSave, App.Title & mcsDefaultRecipientList, "Recipient List")
    
    If sContents & "" <> "" Then
        Call SaveContentsToFile(sContents, Me.txtVotingOptions)
    End If




Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub




Private Sub cmdSaveRecipientsToFile_Click()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : cmdSaveRecipientsToFile_Click
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for cmdSaveRecipientsToFile_Click and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "cmdSaveRecipientsToFile_Click"


    Dim sContents As String
    sContents = GetContentsFileName(ecdlShowSave, App.Title & mcsDefaultRecipientList, "Voting Options")
    
    If sContents & "" <> "" Then
        Call SaveContentsToFile(sContents, Me.cboContacts)
    End If


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

Private Sub cmdSend_Click()

            '__________________________________________________
            ' Scope  : Private
            ' Type   : Sub
            ' Name   : cmdSend_Click
            ' Params :
            ' Returns: Nothing
            ' Desc   : The Sub uses parameters  for cmdSend_Click and returns Nothing.
            '__________________________________________________
            ' History
            ' CDK: 20010105: Added Error Trapping & Comments
            '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "cmdSend_Click"

    Dim oHourglass As clsHourglass
    Dim fEmailSend As Boolean

    Set oHourglass = New clsHourglass

'Validate the required fields and send the e-mail if no errors
    Call ValidateData
    If fErrorFlag = False Then
        fEmailSend = SendEMail
        If fEmailSend = False Then
            MsgBox "The E-Mail could not be sent to the specified recipient", vbInformation + vbOKOnly, "Send failed..."
        Else
            MsgBox "E-Mail has been sent to '" & cboContacts & "'", vbInformation + vbOKOnly, "Send successful..."
        End If
    End If


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    Set oHourglass = Nothing
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

Private Sub ValidateData()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : ValidateData
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for ValidateData and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "ValidateData"


'Validate recipient
    If Trim(cboContacts) = "" Then
        fErrorFlag = True
        MsgBox "You must select a recipient to continue!", vbInformation + vbOKOnly, "No recipient specified"
        Exit Sub
    End If
    
'Validate subject
    If Trim(txtSubject) = "" Then
        fErrorFlag = True
        MsgBox "You must input a subject to continue!", vbInformation + vbOKOnly, "No recipient specified"
        Exit Sub
    End If

'Validate message
    If Trim(txtMessage) = "" Then
        fErrorFlag = True
        MsgBox "You must input message text to continue!", vbInformation + vbOKOnly, "No message to send"
        Exit Sub
    End If
    
    fErrorFlag = False


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

Private Function SendEMail() As Boolean

          '__________________________________________________
          ' Scope  : Private
          ' Type   : Function
          ' Name   : SendEMail
          ' Params :
          ' Returns: Boolean
          ' Desc   : The Function uses parameters  for SendEMail and returns Boolean.
          '__________________________________________________
          ' History
          ' CDK: 20010105: Added Error Trapping & Comments
          '__________________________________________________
    Const csProcName As String = "SendEMail"

    Dim oMailItem As MailItem

    On Error GoTo Proc_Err
    
    
    SendEMail = False
    
'Use the outlook object to create an instance of the MailItem class
    Set oMailItem = moOutlookApplication.CreateItem(olMailItem)
        
'This section sets the properties of the message that are executed
'using the 'Send' method of the MailItem class.
'I have only set some of the obvious properties, but with the
'object exposed you can set attachments, CC's etc
    oMailItem.Importance = olImportanceNormal
    oMailItem.FlagStatus = olNoFlag
    oMailItem.To = cboContacts
    oMailItem.Subject = txtSubject
    oMailItem.Body = txtMessage
    If Me.chkUseVotingOptions = vbChecked _
        And Me.txtVotingOptions & "" <> "" Then
        oMailItem.VotingOptions = Me.txtVotingOptions
    End If
    oMailItem.Send
        
    SendEMail = True
    
    
Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    Set oMailItem = Nothing
    On Error GoTo 0
    Return

Proc_Err:
    
    SendEMail = False

    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

    
    
End Function

Private Function CentreForm(Formname As Form)

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Function
    ' Name   : CentreForm
    ' Params :
    '          Formname As Form
    ' Returns: Nothing
    ' Desc   : The Function uses parameters Formname As Form for CentreForm and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "CentreForm"


'Determine the centre of the screen
    Formname.Move (Screen.Width - Formname.Width) / 2, ((Screen.Height - Formname.Height) / 2)


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Function

Private Sub Search(strAddressBook As String)

       '__________________________________________________
       ' Scope  : Private
       ' Type   : Sub
       ' Name   : Search
       ' Params :
       '          strAddressBook As String
       ' Returns: Nothing
       ' Desc   : The Sub uses parameters strAddressBook As String for Search and returns Nothing.
       '__________________________________________________
       ' History
       ' CDK: 20010105: Added Error Trapping & Comments
       '__________________________________________________


    Const csProcName As String = "Search"

    Dim MyAddressList       As Outlook.AddressList
    Dim MyAddressEntries    As Outlook.AddressEntries
    Dim Index               As Integer
    Dim I                   As Integer
    Dim oHourglass As clsHourglass
    

    On Error GoTo Proc_Err

    Set oHourglass = New clsHourglass
    

    'Open an obj-var over the address book name passed to this sub
    Set MyAddressList = moOutlookNamespace.AddressLists(strAddressBook)
    Set MyAddressEntries = MyAddressList.AddressEntries
    
    'Clear list and add each name in the address book to the list
    cboContacts.Clear
    For I = 1 To MyAddressList.AddressEntries.Count
        lblStatus.Caption = "Retrieving recipient '" & MyAddressList.AddressEntries.Item(I) & "'"
        lblStatus.Refresh
        Me.cboContacts.AddItem MyAddressList.AddressEntries.Item(I)
    Next
    
    lblStatus.Caption = ""
    
Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    Set oHourglass = Nothing
    Set MyAddressList = Nothing
    Set MyAddressEntries = Nothing
        
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly
    
    
End Sub

Private Sub Form_Load()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : Form_Load
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for Form_Load and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "Form_Load"

    
'Test the values of Outlook and initialise
    If moOutlookApplication Is Nothing Then
        If InitializeOutlook = False Then
            MsgBox "Unable to initialize Outlook Application " _
            & "or NameSpace object variables!" & Chr$(13) & "Application terminating...", vbCritical + vbOKOnly, "Could not initialise..."
            End
        End If
    End If
    
'Retrieve info while invisible
    Call RetrieveAddressBooks
    
'Set the form start up properties
    Call SetFormAttributes


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : Form_Unload
    ' Params :
    '          Cancel As Integer
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters Cancel As Integer for Form_Unload and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error Resume Next
    

    
'Release obj-vars and exit
    Set moOutlookApplication = Nothing
    Set moOutlookNamespace = Nothing
    
    
End Sub


Private Sub cboAddressBooks_Click()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : cboAddressBooks_Click
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for cboAddressBooks_Click and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "cboAddressBooks_Click"


'If the user has selected an item, then search through the contents and
'populate the recipients combo
    If cboAddressBooks <> "" Then
        Call Search(cboAddressBooks)
    Else
        cboContacts.Clear
    End If


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

Private Sub RetrieveAddressBooks()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : RetrieveAddressBooks
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for RetrieveAddressBooks and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________


Dim Index               As Integer
    
    On Error GoTo SearchError
    
'Add a blank item at the top of the list
    cboAddressBooks.AddItem ""
    
'Retrieve the current address books available in Outlook
    For Index = 1 To moOutlookNamespace.AddressLists.Count
        lblStatus.Caption = "Retrieving possible address books, please wait..." & moOutlookNamespace.AddressLists.Item(Index)
        lblStatus.Refresh
        cboAddressBooks.AddItem moOutlookNamespace.AddressLists.Item(Index)
    Next
    
    lblStatus.Caption = ""
    
    Exit Sub
    
SearchError:
    lblStatus.Caption = ""
    MsgBox "Error (" & Err.Description & ") has occurred" & Chr$(13) & "Search cancelled", vbCritical + vbOKOnly, "Error number " & Err.Number

End Sub


Private Sub SetFormAttributes()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : SetFormAttributes
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for SetFormAttributes and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "SetFormAttributes"


'The start up properties of the form
    Call CentreForm(Me)
    Me.Visible = True
    'lstContacts.Enabled = False
    cboAddressBooks.SetFocus


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub



Private Function GetContentsFileName(Optional eMode As eCommonDialogConst = ecdlShowOpen, Optional sDefaultContents As String = "", Optional sDefaultDialogTitle As String = mcsDefaultDialogTitle) As String

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Function
    ' Name   : GetContentsFileName
    ' Params :
    '          Optional eMode As eCommonDialogConst = ecdlShowOpen
    '          Optional sDefaultContents As String = ""
    '          Optional sDefaultDialogTitle As String = mcsDefaultDialogTitle
    ' Returns: String
    ' Desc   : The Function uses parameters Optional eMode As eCommonDialogConst = ecdlShowOpen, Optional sDefaultContents As String = "" and Optional sDefaultDialogTitle As String = mcsDefaultDialogTitle for GetContentsFileName and returns String.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________


    On Error GoTo Proc_Exit
    Dim sFile As String
    Dim fFileExists As Boolean
    Dim lResponse As Long
    
    cdlgMain.DialogTitle = sDefaultDialogTitle
    cdlgMain.InitDir = App.Path
    cdlgMain.DefaultExt = mcsDefaultExtention
    cdlgMain.FileName = sDefaultContents
    cdlgMain.Filter = mcsFilterForCommonDialog
    cdlgMain.CancelError = True
    cdlgMain.Flags = cdlOFNHideReadOnly
    cdlgMain.Action = eMode
    
    
    
    Select Case eMode
        Case eCommonDialogConst.ecdlShowSave
            If Len(cdlgMain.FileName) = 0 Then
                'no file typed in
                Err.Raise Number:=4001, Description:="no file selected"
            
            Else
                sFile = cdlgMain.FileName
                fFileExists = CBool(Len(Dir(sFile)))
                If fFileExists Then
                    lResponse = MsgBox("File already exists, overwrite existing file?", vbQuestion + vbYesNoCancel, App.Title)
                    If lResponse <> vbYes Then
                        sFile = ""
                    End If
                End If
            End If
        Case eCommonDialogConst.ecdlShowOpen
            If Len(cdlgMain.FileName) > 0 And Len(Dir(cdlgMain.FileName)) > 0 Then
                sFile = cdlgMain.FileName
            Else
                Err.Raise Number:=4001, Description:="Cannot find file"
            End If
            
    End Select
    
Proc_Exit:

    GetContentsFileName = sFile

End Function



Private Sub SaveContentsToFile(sFileName As String, oCtl As Control)

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : SaveContentsToFile
    ' Params :
    '          sFileName As String
    '          oCtl As Control
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters sFileName As String and oCtl As Control for SaveContentsToFile and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "SaveContentsToFile"



    Dim ofs As clsFs
    Dim oTs As TextStream
    
    Set ofs = New clsFs
    Set oTs = ofs.CreateTextFile(sFileName, True, False)
    oTs.Write oCtl.Text





Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        
    Set oTs = Nothing
    Set ofs = Nothing
    

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub



Private Sub LoadContentsFromFile(sContentsToLoad As String, oCtlToLoad As Control)

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : LoadContentsFromFile
    ' Params :
    '          sContentsToLoad As String
    '          oCtlToLoad As Control
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters sContentsToLoad As String and oCtlToLoad As Control for LoadContentsFromFile and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "LoadContentsFromFile"


    Dim ofs As clsFs
    Dim oTs As TextStream

    Set ofs = New clsFs


    'see if the file exists for the default information
    If ofs.FileExists(sContentsToLoad) Then
        Set oTs = ofs.GetFile(sContentsToLoad).OpenAsTextStream
        oCtlToLoad = oTs.ReadAll
    End If





Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        Set ofs = Nothing
    Set oTs = Nothing

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub



Private Sub cmdClear_Click()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : cmdClear_Click
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for cmdClear_Click and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "cmdClear_Click"

    
    Dim oCtl As Control

    For Each oCtl In Me.Controls

        If TypeOf oCtl Is TextBox Then
            oCtl.Text = ""
        ElseIf TypeOf oCtl Is ListBox Then
            oCtl.Clear
        ElseIf TypeOf oCtl Is ComboBox Then
            oCtl.Text = ""
            
        ElseIf TypeOf oCtl Is CheckBox Then
            oCtl.Value = vbUnchecked
        
        End If

    Next




Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        Set oCtl = Nothing

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmEMail->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub





