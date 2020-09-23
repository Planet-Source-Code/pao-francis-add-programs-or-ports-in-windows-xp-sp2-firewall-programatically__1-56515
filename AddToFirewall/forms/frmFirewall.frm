VERSION 5.00
Begin VB.Form frmFirewall 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exceptions - Windows XP SP2 Firewall"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFirewall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAuthor 
      Caption         =   "About..."
      Height          =   315
      Left            =   90
      TabIndex        =   32
      Top             =   6510
      Width           =   1185
   End
   Begin VB.TextBox txtDescSPG 
      Height          =   315
      Left            =   1530
      TabIndex        =   3
      Top             =   1620
      Width           =   3675
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4020
      TabIndex        =   30
      Top             =   6510
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2820
      TabIndex        =   29
      Top             =   6510
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   1500
      TabIndex        =   26
      Top             =   5010
      Width           =   3555
      Begin VB.OptionButton optAllIPSP 
         Caption         =   "All IP addresses"
         Height          =   195
         Left            =   30
         TabIndex        =   13
         Top             =   60
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optLocSubSP 
         Caption         =   "Local Subnet Only"
         Height          =   195
         Left            =   1710
         TabIndex        =   14
         Top             =   60
         Width           =   1905
      End
      Begin VB.OptionButton optCustSP 
         Caption         =   "Custom List:"
         Height          =   195
         Left            =   30
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCustListSP 
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   600
         Width           =   3435
      End
      Begin VB.Label Label9 
         Caption         =   "Example: 192.168.114.201 OR 192.168.114.201/255.255.255.0"
         Height          =   405
         Left            =   60
         TabIndex        =   27
         Top             =   930
         Width           =   3435
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1530
      TabIndex        =   24
      Top             =   4710
      Width           =   3495
      Begin VB.OptionButton optUDP 
         Caption         =   "UDP"
         Height          =   195
         Left            =   1680
         TabIndex        =   12
         Top             =   30
         Width           =   855
      End
      Begin VB.OptionButton optTCP 
         Caption         =   "TCP"
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   30
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   1530
      MaxLength       =   5
      TabIndex        =   10
      Top             =   4260
      Width           =   1305
   End
   Begin VB.TextBox txtDescSP 
      Height          =   315
      Left            =   1530
      TabIndex        =   9
      Top             =   3900
      Width           =   3675
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Specify a port:"
      Height          =   225
      Left            =   180
      TabIndex        =   31
      Top             =   3570
      Width           =   4905
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   1530
      TabIndex        =   20
      Top             =   1980
      Width           =   3555
      Begin VB.TextBox txtCustListSPG 
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   600
         Width           =   3435
      End
      Begin VB.OptionButton optCustSPG 
         Caption         =   "Custom List:"
         Height          =   195
         Left            =   30
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optLocSubSPG 
         Caption         =   "Local Subnet Only"
         Height          =   195
         Left            =   1710
         TabIndex        =   5
         Top             =   60
         Width           =   1905
      End
      Begin VB.OptionButton optAllIPSPG 
         Caption         =   "All IP addresses"
         Height          =   195
         Left            =   30
         TabIndex        =   4
         Top             =   60
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Example: 192.168.114.201 OR 192.168.114.201/255.255.255.0"
         Height          =   405
         Left            =   60
         TabIndex        =   25
         Top             =   930
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   315
      Left            =   3990
      TabIndex        =   2
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1260
      Width           =   2415
   End
   Begin VB.OptionButton OptProg 
      Caption         =   "Select a program:"
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   930
      Value           =   -1  'True
      Width           =   5025
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Scope:"
      Height          =   195
      Left            =   870
      TabIndex        =   28
      Top             =   5040
      Width           =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Protocol:"
      Height          =   195
      Left            =   735
      TabIndex        =   23
      Top             =   4740
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Port Number:"
      Height          =   195
      Left            =   345
      TabIndex        =   22
      Top             =   4350
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   465
      TabIndex        =   21
      Top             =   3990
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   90
      Picture         =   "frmFirewall.frx":12B2
      Top             =   180
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Scope:"
      Height          =   195
      Left            =   900
      TabIndex        =   19
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   465
      TabIndex        =   18
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
      Height          =   195
      Left            =   660
      TabIndex        =   17
      Top             =   1350
      Width           =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You can allow a program to automatically communicate  through the firewall, or you can manually open a port."
      ForeColor       =   &H8000000E&
      Height          =   645
      Left            =   720
      TabIndex        =   8
      Top             =   180
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   705
      Left            =   660
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmFirewall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- Version 1.0.0
'---- 10.04.2004 4:07PM Philippine Manila Time
'---- this is a vb6 version of howTo add programs/port programatically in windows sp2 firewall.
'---- VB.net/C#.net source found @ http://msdn.microsoft.com/security/productinfo/XPSP2/networkprotection/firewall_devimp.aspx
'---- I was inspired by the source code @ http://weblogs.asp.net/sjoshi/archive/2004/07/07/175309.aspx
'---- so I made my own frontend out of it. cheers to the original creator!

'---- Notes:
'---- Listen! You need to add 'NetFwTypeLib' & 'NetCon 1.0 Type Library' in the Reference Tab.
'---- You may distribute this code, make it your own. I dont care, Im only here to help & motivate!

'---- Windows Firewall is a copyright of Microsoft Corporation.
'---- All sources codes etc... is trademark of whomever who made it.
'---- the firewall icon is a copyright of Microsoft Corporation.

'---- Pls. have your VB Updated to Service Pack 6!
'---- This app assumes that the current OS has ALREADY SP2 Installed! all WINXP!

'---- All the rest is up to you, cheers!

Option Explicit

Private Enum en_OptTypeEnable
    enProggy
    enPort
End Enum

'---- file open in api a.k.a. common dialog api
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private OFName As OPENFILENAME

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private PROGSCOPE As NET_FW_SCOPE_
Private PORTSCOPE As NET_FW_SCOPE_
Private PROTOCOL  As NET_FW_IP_PROTOCOL_

Private Sub cmdAuthor_Click()
    MsgBox "Very basic app of howTo add applications etc.. in the winxp sp2 firewall." & vbCrLf & _
           "feel free to msg me in pscode.com :-)", vbInformation + vbOKOnly, "'bout me..."
End Sub

Private Sub cmdBrowse_Click()
    If GetOpenFileName(OFName) Then txtFilename.Text = Trim(OFName.lpstrFile)
End Sub

Private Sub cmdCancel_Click()
    Unload Me: End
End Sub

Private Sub cmdOK_Click()
Dim objFirewall As INetFwMgr
Dim objAuthApp  As INetFwAuthorizedApplication
Dim objPort     As INetFwOpenPort
Dim objProfile  As Object

Dim sMsg        As String
Dim sSuccessMsg As String

On Error GoTo Err_cmdOK

    Set objFirewall = CreateObject("HNetCfg.FwMgr")
    Set objAuthApp = CreateObject("HNetCfg.FwAuthorizedApplication")
    Set objProfile = objFirewall.LocalPolicy.CurrentProfile
    Set objPort = CreateObject("HNetCfg.FwOpenPort")

    If OptProg.Value Then '---- Main Add Program to winxp sp2 firewall is here!
        '---- Validation
        If Not ValidateFireCtrls(enProggy, sMsg) Then Err.Description = sMsg: GoTo Err_cmdOK
        
        With objAuthApp
            .Name = txtDescSPG.Text
            .ProcessImageFileName = txtFilename.Text
            .Enabled = True  '---- place a checkmark in win. firewall
            .IpVersion = NET_FW_IP_VERSION_ANY
            .Scope = PROGSCOPE
        End With
        
        objProfile.AuthorizedApplications.Add objAuthApp
        sSuccessMsg = "The following has been added..." & vbCrLf & _
                      "App. name : " & txtDescSPG.Text & vbCrLf & _
                      "App. path : " & txtFilename.Text
                      
    ElseIf optPort.Value Then '---- Main Add Port to winxp sp2 firewall is here!
        
        '---- Validation
        If Not ValidateFireCtrls(enPort, sMsg) Then Err.Description = sMsg: GoTo Err_cmdOK
        
        With objPort
            .Name = txtDescSP.Text
            .Port = txtPort.Text
            .Scope = PORTSCOPE
            .PROTOCOL = PROTOCOL
            .Enabled = True '---- place a checkmark in win. firewall
            objProfile.GloballyOpenPorts.Add objPort
        End With
        
        sSuccessMsg = "The following has been added..." & vbCrLf & _
                      "App. name : " & txtDescSP.Text & vbCrLf & _
                      "App. port : " & txtPort.Text
        
    End If
    
    MsgBox sSuccessMsg, vbInformation + vbOKOnly, "Pls. check your windows firewall."
    
    GoTo Exiting
    
Err_cmdOK:

    MsgBox "Error! " & Err.Description, vbCritical + vbOKOnly
    
Exiting:
    Set objFirewall = Nothing
    Set objAuthApp = Nothing
    Set objProfile = Nothing
    Set objPort = Nothing
    
End Sub

Private Sub Form_Load()
    
    '---- Initialize Api Getfilename
    InitFOpenSettings
    
    '---- Default Settings
    OptEnable enProggy
    PROGSCOPE = NET_FW_SCOPE_ALL
    PROTOCOL = NET_FW_IP_PROTOCOL_TCP
    PORTSCOPE = NET_FW_SCOPE_ALL
    
End Sub

Private Sub optAllIPSP_Click()
    PORTSCOPE = NET_FW_SCOPE_ALL
End Sub

Private Sub optAllIPSPG_Click()
    PROGSCOPE = NET_FW_SCOPE_ALL
End Sub

Private Sub optCustSP_Click()
    PORTSCOPE = NET_FW_SCOPE_CUSTOM
    txtCustListSP.SetFocus
End Sub

Private Sub optCustSPG_Click()
    PROGSCOPE = NET_FW_SCOPE_CUSTOM
    txtCustListSPG.SetFocus
End Sub

Private Sub optLocSubSP_Click()
    PORTSCOPE = NET_FW_SCOPE_LOCAL_SUBNET
End Sub

Private Sub optLocSubSPG_Click()
    PROGSCOPE = NET_FW_SCOPE_LOCAL_SUBNET
End Sub

Private Sub optPort_Click()
    If optPort.Value Then OptEnable enPort
End Sub

Private Sub OptProg_Click()
    If OptProg.Value Then OptEnable enProggy
End Sub

Private Sub optTCP_Click()
    PROTOCOL = NET_FW_IP_PROTOCOL_TCP
End Sub

Private Sub optUDP_Click()
    PROTOCOL = NET_FW_IP_PROTOCOL_UDP
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPort_Validate(Cancel As Boolean)
    If Val(txtPort.Text) > 65000 Then
        MsgBox "Port number is invalid!", vbCritical + vbOKOnly, "Error!"
        txtPort.Text = ""
    End If
End Sub

Private Sub InitFOpenSettings()
    With OFName
        .lStructSize = Len(OFName)
        .hwndOwner = Me.hWnd
        .hInstance = App.hInstance
        .lpstrFilter = "Applications (*.exe;*.com;*.icd)" + Chr$(0) + "*.exe" + Chr$(0) + Chr$(0) + "*.com" + Chr$(0) + Chr$(0) + "*.icd" + Chr$(0)
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = App.Path
        .lpstrTitle = "Browse"
        .flags = 0
    End With
End Sub

Private Sub OptEnable(ByVal pType As en_OptTypeEnable)

    Select Case pType
        
        Case enProggy
            
            '---- Enable [choose programs]
            txtFilename.Enabled = True
            txtFilename.BackColor = vbWhite
            txtDescSPG.Enabled = True
            txtDescSPG.BackColor = vbWhite
            optAllIPSPG.Enabled = True
            optLocSubSPG.Enabled = True
            optCustSPG.Enabled = True
            txtCustListSPG.Enabled = True
            txtCustListSPG.BackColor = vbWhite
            
            '---- Disable [spes ports]
            txtDescSP.Enabled = False
            txtDescSP.BackColor = &HC0C0C0
            txtPort.Enabled = False
            txtPort.BackColor = &HC0C0C0
            optTCP.Enabled = False
            optUDP.Enabled = False
            
            optAllIPSP.Enabled = False
            optLocSubSP.Enabled = False
            optCustSP.Enabled = False
            txtCustListSP.Enabled = False
            txtCustListSP.BackColor = &HC0C0C0
            
        Case enPort
            
            '---- Disable [choose programs]
            txtFilename.Enabled = False
            txtFilename.BackColor = &HC0C0C0
            txtDescSPG.Enabled = False
            txtDescSPG.BackColor = &HC0C0C0
            optAllIPSPG.Enabled = False
            optLocSubSPG.Enabled = False
            optCustSPG.Enabled = False
            txtCustListSPG.Enabled = False
            txtCustListSPG.BackColor = &HC0C0C0
            
            '---- Enable [spes ports]
            txtDescSP.Enabled = True
            txtDescSP.BackColor = vbWhite
            txtPort.Enabled = True
            txtPort.BackColor = vbWhite
            optTCP.Enabled = True
            optUDP.Enabled = True
            
            optAllIPSP.Enabled = True
            optLocSubSP.Enabled = True
            optCustSP.Enabled = True
            txtCustListSP.Enabled = True
            txtCustListSP.BackColor = vbWhite
            
    End Select
    
End Sub

Private Function ValidateFireCtrls(ByVal pType As en_OptTypeEnable, _
                                   ByRef pErrMsg As String) As Boolean
On Error GoTo Err_ValidateFireCtrls

    ValidateFireCtrls = True
    
    Select Case pType
        Case enProggy
            If Trim(txtFilename.Text) = "" Then
                pErrMsg = "You must place a valid file path."
                txtFilename.SetFocus
                GoTo Err_ValidateFireCtrls
            End If
            
            If Trim(txtDescSPG.Text) = "" Then
                pErrMsg = "Pls. place a description."
                txtDescSPG.SetFocus
                GoTo Err_ValidateFireCtrls
            End If
            
            If optCustSPG.Value Then
                If Trim(txtCustListSPG.Text) = "" Then
                    pErrMsg = "Pls. place a custom list."
                    txtCustListSPG.SetFocus
                    GoTo Err_ValidateFireCtrls
                End If
            End If
            
        Case enPort
            If Trim(txtDescSP.Text) = "" Then
                pErrMsg = "Pls. place a description."
                txtDescSP.SetFocus
                GoTo Err_ValidateFireCtrls
            End If
            
            If Trim(txtPort.Text) = "" Then
                pErrMsg = "Pls. place a port."
                txtPort.SetFocus
                GoTo Err_ValidateFireCtrls
            End If
            
            If optCustSPG.Value Then
                If Trim(txtCustListSP.Text) = "" Then
                    pErrMsg = "Pls. place a custom list."
                    txtCustListSP.SetFocus
                    GoTo Err_ValidateFireCtrls
                End If
            End If
            
    End Select
    
    Exit Function
    
Err_ValidateFireCtrls:
    ValidateFireCtrls = False
    
End Function

