VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODBUS"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   19860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   90
      ScaleHeight     =   7065
      ScaleWidth      =   17655
      TabIndex        =   29
      Top             =   1440
      Width           =   17685
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   18030
      TabIndex        =   28
      Text            =   "1"
      Top             =   2430
      Width           =   1455
   End
   Begin VB.TextBox Text_Display 
      Alignment       =   1  'Right Justify
      CausesValidation=   0   'False
      Height          =   285
      Index           =   0
      Left            =   18210
      TabIndex        =   26
      Text            =   "0"
      Top             =   1890
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Reset_Comm_Status 
      Appearance      =   0  'Flat
      Caption         =   "Reset Status"
      Height          =   290
      Left            =   9720
      TabIndex        =   25
      Top             =   8880
      Width           =   1365
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   8820
      Width           =   19860
      _ExtentX        =   35031
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3793
            MinWidth        =   3793
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2911
            MinWidth        =   2911
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3351
            MinWidth        =   3351
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   0
      TabIndex        =   24
      Top             =   8640
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Scrolling       =   1
   End
   Begin VB.CommandButton Reset_Text 
      Caption         =   "Reset"
      Height          =   465
      Left            =   18030
      TabIndex        =   23
      Top             =   6300
      Width           =   1365
   End
   Begin VB.Frame Frame_Funtion_Code 
      Caption         =   "Select Funtion Code"
      Height          =   1215
      Left            =   7710
      TabIndex        =   21
      Top             =   210
      Width           =   1935
      Begin VB.ComboBox cmbFunctioCode 
         Height          =   315
         Left            =   60
         TabIndex        =   22
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Label_CRC_Status 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   690
         Width           =   1725
      End
   End
   Begin VB.Frame Time_Config 
      Caption         =   "Time Configuration (mS)"
      Height          =   1215
      Left            =   5400
      TabIndex        =   13
      Top             =   210
      Width           =   2295
      Begin VB.TextBox Delay_BWPolls 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   17
         Text            =   "1000"
         Top             =   580
         Width           =   465
      End
      Begin VB.TextBox TimeOut_SlaveResponse 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1890
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "10"
         Top             =   280
         Width           =   315
      End
      Begin VB.Label Label5 
         Caption         =   "Delay Between Polls"
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Slave Response Timeout"
         Height          =   225
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   1845
      End
   End
   Begin VB.Frame Frame_BaudRate 
      Caption         =   "BaudRate"
      Height          =   1215
      Left            =   3390
      TabIndex        =   10
      Top             =   210
      Width           =   1005
      Begin VB.OptionButton BaudRate 
         Caption         =   "11500"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   32
         Top             =   900
         Width           =   795
      End
      Begin VB.OptionButton BaudRate 
         Caption         =   "9600"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   12
         Top             =   600
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton BaudRate 
         Caption         =   "4800"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   330
         Width           =   675
      End
   End
   Begin VB.CommandButton Start 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   465
      Left            =   18030
      TabIndex        =   1
      Top             =   240
      Width           =   1365
   End
   Begin VB.Frame Frame_Comm_Mode 
      Caption         =   "Mode"
      Height          =   1215
      Left            =   4410
      TabIndex        =   7
      Top             =   210
      Width           =   975
      Begin VB.OptionButton Command_Mode_Polling 
         Caption         =   "Polling"
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Top             =   600
         Width           =   825
      End
      Begin VB.OptionButton Command_Mode_Single 
         Caption         =   "Single"
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   17820
      Top             =   930
   End
   Begin VB.Timer Timer_Receive 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   18900
      Top             =   900
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   18240
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame_Port_Setting 
      Caption         =   "Port Setting"
      Height          =   1215
      Left            =   2280
      TabIndex        =   6
      Top             =   210
      Width           =   1095
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   150
         TabIndex        =   31
         Text            =   "1"
         Top             =   450
         Width           =   825
      End
   End
   Begin VB.Frame Frame_Device_Specifications 
      Caption         =   "Device Specifications"
      Height          =   1215
      Left            =   90
      TabIndex        =   2
      Top             =   210
      Width           =   2175
      Begin VB.TextBox txtRegisterCount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "1"
         Top             =   810
         Width           =   435
      End
      Begin VB.TextBox txtStartRegister 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "1"
         Top             =   495
         Width           =   675
      End
      Begin VB.TextBox txtDeviceID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "1"
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Device ID"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Register Address"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   570
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Register Count"
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   870
         Width           =   1155
      End
   End
   Begin VB.Label Label_Text_Display 
      Caption         =   "Label6"
      Height          =   255
      Index           =   0
      Left            =   18150
      TabIndex        =   30
      Top             =   1590
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim g_Buffer As String
    Dim g_Port As Byte
    Dim g_CommandMode As Byte
    Dim g_BaudRate As Long
    Dim a_sCommand As String
    Dim g_ReceiveBuffer(0 To 255) As Byte
    Dim g_ReceiveIndex As Byte
    Dim g_PollesCount As Long
    Dim g_ResponseCount As Long
    Dim g_RegisterAddress As Long
    Dim g_RgisterCount As Integer
    Dim g_ProgressBarIndex As Byte
    Dim g_ProgressBarMaxValue As Byte

Private Sub BaudRate_Click(Index As Integer)
    Dim l_str As String
        
    l_str = GetBaudrateFromIndex(Index)
    g_BaudRate = Val(l_str)
    StatusBar1.Panels(2).Text = "Baudrate:  " & Val(g_BaudRate)
    Connect_Click
End Sub

Private Function GetBaudrateFromIndex(a_Index As Integer) As Long
    Dim l_retVal As Long
    
    l_retVal = 0
    
    Select Case a_Index
        Case 0
            l_retVal = 4800
        Case 1
            l_retVal = 9600
        Case 2
            l_retVal = 115200
    End Select
    
    GetBaudrateFromIndex = l_retVal
End Function

Private Sub Command_Mode_Polling_Click()
    Timer1.Interval = Val(Delay_BWPolls)
    Timer1.Enabled = True
    StatusBar1.Panels(3).Text = "Request Mode:  Polling"
End Sub

Private Sub Command_Mode_Single_Click()
    Timer1.Interval = 0
    Timer1.Enabled = False
    StatusBar1.Panels(3).Text = "Request Mode:  Single"
End Sub

Private Sub Connect_Click()
    With MSComm1
        If .PortOpen Then .PortOpen = False
        .CommPort = Val(txtPort)
        .Settings = g_BaudRate & ",N,8,1"
        .Handshaking = comNone
        '.RTSEnable = True
        .RThreshold = 1
        .SThreshold = 1
        .InputLen = 1
        '.InputMode = comInputModeBinary
        .PortOpen = True
    End With
        
    g_Buffer = MSComm1.Input
    g_Buffer = ""
End Sub

Private Sub Delay_BWPolls_Validate(Cancel As Boolean)
    If Val(Delay_BWPolls) < (Val(txtRegisterCount) * 6) Then
        MsgBox "Invalid Interval Between Polls Please Enter Greater Then " & ((Val(txtRegisterCount) * 6) - 1)
        Cancel = True
    End If
End Sub

Private Sub Picture1_DblClick()
    Picture1.Cls
End Sub

Private Sub Port_Click(Index As Integer)
    g_Port = Index
    StatusBar1.Panels(1).Text = "Communication Port:  COM" & Val(g_Port)
End Sub


Private Sub Command_Mode_Click(Index As Integer)
    If Index = 1 Then
        Timer1.Enabled = (Frame_Comm_Mode.Index = 1)    '.Value = vbChecked)
    End If
End Sub


Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "Application is Already Running !!!", vbCritical, App.Title
        End
    End If
    ''''ResetDisplay
    Command_Mode_Single_Click
    'Command_Mode_Polling_Click
    Timer1.Enabled = False
    
    With cmbFunctioCode
        .AddItem ("03: Read Holding Register")
        .AddItem ("04: Read Input Register")
        .AddItem ("06: Write a Register")
        .AddItem ("16: Write Multiple Registers")
        
        .ListIndex = 0
    End With
    
    Timer_Receive.Interval = Val(TimeOut_SlaveResponse)
    g_Port = 1
    g_BaudRate = 9600
    g_ProgressBarIndex = 0
    g_ProgressBarMaxValue = 100
    ProgressBar1.Value = g_ProgressBarIndex
    ProgressBar1.Max = g_ProgressBarMaxValue
    ResetStatusBar
   
End Sub

Private Sub MSComm1_OnComm()
    If MSComm1.CommEvent <> comEvReceive Then Exit Sub

    g_Buffer = MSComm1.Input
    g_ReceiveBuffer(g_ReceiveIndex) = Asc(g_Buffer)
    
    Picture1.Print PadLeft(Hex(g_ReceiveBuffer(g_ReceiveIndex)), 2, "0") & " ";
    g_ReceiveIndex = g_ReceiveIndex + 1
    
    g_ProgressBarIndex = g_ProgressBarIndex + 1
    If g_ProgressBarIndex > g_ProgressBarMaxValue Then g_ProgressBarIndex = g_ProgressBarMaxValue
    ProgressBar1.Value = g_ProgressBarIndex
    
End Sub

Private Sub Reset_Counts_Click()
    ResetStatusBar
End Sub

Private Sub Reset_Comm_Status_Click()
    ResetStatusBar
End Sub

Private Sub Reset_Text_Click()
    ''''ResetDisplay
End Sub

Private Sub Start_Click()
    If Not MSComm1.PortOpen Then
        Connect_Click   'connect
    End If
    
''''    If Start.Caption = "Send" Then
''''        LockParameter
        Send_Command
''''        Timer1.Interval = Val(Delay_BWPolls)
''''        Timer1.Enabled = True
''''        Start.Caption = "Stop"
''''    ElseIf Start.Caption = "Send" Then
''''        UnLockParameter
''''        Send_Command
''''        Timer1.Enabled = False
''''    Else
''''        UnLockParameter
''''        Start.Caption = "Start"
''''        Timer1.Enabled = False
''''    End If
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
    ResetStatusBar
End Sub

Private Sub txtDeviceID_Validate(Cancel As Boolean)
''''    If Val(txtDeviceID) > 247 Or Val(txtDeviceID) = 0 Then
''''        MsgBox "Invalid Device ID Please Enter Less Then 1 to 247"
''''        Cancel = True
''''    End If
End Sub

Private Sub txtRegisterCount_Validate(Cancel As Boolean)
    If Val(txtRegisterCount) > 127 Or Val(txtRegisterCount) = 0 Then
        MsgBox "Invalid Point to Read Please Enter Less Then 1 to 127"
        Cancel = True
        Exit Sub
    End If
    If (Val(txtRegisterCount) * 6) > Val(Delay_BWPolls) Then
        MsgBox "Invalid Interval Between Polls Please Enter Value Greater Then " & ((Val(txtRegisterCount) * 6) - 1)
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub txtStartRegister_Validate(Cancel As Boolean)
    If Val(txtStartRegister) = 0 Then
        MsgBox "Invalid Register_Address Please Enter Greater Then 1"
        Cancel = True
    End If
End Sub

Private Sub TimeOut_SlaveResponse_Validate(Cancel As Boolean)
    If Val(TimeOut_SlaveResponse) < 10 Then
        MsgBox "Invalid Slave Response Time Please Enter Greater Then 10"
        Cancel = True
    End If
End Sub

Private Sub Timer_Receive_Timer()
    Timer_Receive.Enabled = False
    
    Dim l_cNoOfRegister As Byte
    Dim l_iCRC As Long
    Dim l_iCRCTemp As Integer
    Dim l_lTempCheck As Long
    Dim l_cTemp As Byte
    Dim l_sReceive As String
    
    l_iCRC = 0
    Label_CRC_Status = ""
    If Hex(Asc(Mid(a_sCommand, 1, 1))) = Hex(g_ReceiveBuffer(0)) Then           'Device ID
        l_cTemp = 4 + g_ReceiveBuffer(2)
        l_iCRC = g_ReceiveBuffer(l_cTemp)
        l_iCRC = l_iCRC * 256
        l_iCRC = l_iCRC + g_ReceiveBuffer(l_cTemp - 1)
        l_iCRCTemp = MODBUS_CRC_16_Received(g_ReceiveBuffer, l_cTemp - 1)
        l_lTempCheck = l_iCRC - l_iCRCTemp
        If l_lTempCheck = 0 Or l_lTempCheck = 65536 Then                        'CRC
            If Hex(Asc(Mid(a_sCommand, 2, 1))) = Hex(g_ReceiveBuffer(1)) Then   'Function Code
                l_cNoOfRegister = g_ReceiveBuffer(2) \ 2
                g_ResponseCount = g_ResponseCount + 1
                'Text_Response_Count = g_ResponseCount
                StatusBar1.Panels(5).Text = "Response:  " & Val(g_ResponseCount)
                ''''l_lTempCheck = FillDisplay(g_ReceiveBuffer, g_RegisterAddress, l_cNoOfRegister)
            End If                                                              'Function Code End
        Else
            Label_CRC_Status = Val(l_iCRC - l_iCRCTemp)
            l_iCRC = 0
        End If                                                                  'CRC End
    End If                                                                      'Device ID End

    g_ReceiveIndex = 0
    a_sCommand = ""
    
    For l_cNoOfRegister = 0 To 255 - 1
        g_ReceiveBuffer(l_cNoOfRegister) = 0
    Next
    
    g_ProgressBarIndex = 0
    ProgressBar1.Value = g_ProgressBarIndex
End Sub

Private Sub Timer1_Timer()
    Picture1.Print
    Send_Command
End Sub
Private Function ConvertHexAscii2Hex(aData As String) As String
    Dim i As Integer
    Dim bytData As Byte
    Dim sTemp As String
    
    If Len(aData) Mod 2 = 1 Then
        aData = "0" & aData
    End If
    
    sTemp = ""
    
    For i = 1 To Len(aData) Step 2
        bytData = "&h" & Mid(aData, i, 2)
        sTemp = sTemp & Chr(bytData)
    Next i
    ConvertHexAscii2Hex = sTemp
End Function

Private Sub Send_Command()
    Select Case Mid(cmbFunctioCode.Text, 1, 2)
        Case "03", "04"
            MBReadRegisters
        Case "06"
            MBWriteSingleRegister
        Case "16"
            MBWriteMultiRegisters
    End Select
End Sub

Public Sub MBWriteMultiRegisters()
    Dim l_sCommandBuffer As String
    Dim l_sSendData As String
    Dim l_lTemp As Long
    Dim l_iCRC As Integer
    Dim l_FunctionCode As Byte
        
    Picture1.Print
         
    l_sCommandBuffer = ""
    l_sSendData = ""
    g_Buffer = ""
    
    l_sCommandBuffer = ConvertHexAscii2Hex(Hex(Val(txtDeviceID.Text)))
    
    l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(Val(Mid(cmbFunctioCode.Text, 1, 2))))
    
    l_lTemp = Int(Val(txtStartRegister))
    g_RegisterAddress = l_lTemp
    If l_lTemp > 0 Then l_lTemp = l_lTemp - 1
    
    'Starting Register Address
    l_sCommandBuffer = l_sCommandBuffer & PadLeft(ConvertHexAscii2Hex(Hex(Val(l_lTemp))), 2, Chr(0))

    'Number of registers
    l_sCommandBuffer = l_sCommandBuffer & PadLeft(ConvertHexAscii2Hex(Hex(Val(txtRegisterCount))), 2, Chr(0))
    
    'Byte Count
    l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(2 * Val(txtRegisterCount)))
    
    'Data filling
    l_sCommandBuffer = l_sCommandBuffer & PadLeft(ConvertHexAscii2Hex(Hex(Val(txtData))), 2 * Val(txtRegisterCount), Chr(0))
    
    'CRC Calculation
    l_iCRC = MODBUS_CRC_16(l_sCommandBuffer, Len(l_sCommandBuffer))
    l_sCommandBuffer = l_sCommandBuffer & StrReverse(ConvertHexAscii2Hex(Hex(l_iCRC)))
    
    a_sCommand = l_sCommandBuffer
    
    g_ReceiveIndex = 0
    g_ProgressBarIndex = 0
    g_ProgressBarMaxValue = 5 + (g_RgisterCount * 2)
    ProgressBar1.Value = g_ProgressBarIndex
    ProgressBar1.Max = g_ProgressBarMaxValue
    
    ''''MSComm1.Output = l_sCommandBuffer
    WriteToPort l_sCommandBuffer
    
    g_PollesCount = g_PollesCount + 1
    'Text_Polles_Count = g_PollesCount
    StatusBar1.Panels(4).Text = "Polles:  " & Val(g_PollesCount)
    
    Timer_Receive.Interval = Val(TimeOut_SlaveResponse)
    Timer_Receive.Enabled = True
    
    End Sub

Public Sub MBWriteSingleRegister()
    Dim l_sCommandBuffer As String
    Dim l_sSendData As String
    Dim l_lTemp As Long
    Dim l_iCRC As Integer
    Dim l_FunctionCode As Byte
        
    Picture1.Print
         
    l_sCommandBuffer = ""
    l_sSendData = ""
    g_Buffer = ""
    
    l_sCommandBuffer = ConvertHexAscii2Hex(Hex(Val(txtDeviceID.Text)))
    
    l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(Val(Mid(cmbFunctioCode.Text, 1, 2))))
    
    l_lTemp = Int(Val(txtStartRegister))
    g_RegisterAddress = l_lTemp
    If l_lTemp > 0 Then l_lTemp = l_lTemp - 1
    

    If l_lTemp > 255 Then
        l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(l_lTemp))
    Else
        l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(0))
        l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(l_lTemp))
    End If

    'Data filling
    l_sCommandBuffer = l_sCommandBuffer & PadLeft(ConvertHexAscii2Hex(Hex(Val(txtData))), 2, Chr(0))
    
    l_iCRC = MODBUS_CRC_16(l_sCommandBuffer, 6)
    l_sCommandBuffer = l_sCommandBuffer & StrReverse(ConvertHexAscii2Hex(Hex(l_iCRC)))
    
    a_sCommand = l_sCommandBuffer
    
    g_ReceiveIndex = 0
    g_ProgressBarIndex = 0
    g_ProgressBarMaxValue = 5 + (g_RgisterCount * 2)
    ProgressBar1.Value = g_ProgressBarIndex
    ProgressBar1.Max = g_ProgressBarMaxValue
    
    ''''MSComm1.Output = l_sCommandBuffer
    WriteToPort l_sCommandBuffer
    
    g_PollesCount = g_PollesCount + 1
    'Text_Polles_Count = g_PollesCount
    StatusBar1.Panels(4).Text = "Polles:  " & Val(g_PollesCount)
    
    Timer_Receive.Interval = Val(TimeOut_SlaveResponse)
    Timer_Receive.Enabled = True
    
    End Sub

Public Sub MBReadRegisters()
    Dim l_sCommandBuffer As String
    Dim l_sSendData As String
    Dim l_lTemp As Long
    Dim l_iCRC As Integer
    Dim l_FunctionCode As Byte
        
    Picture1.Print
         
    l_sCommandBuffer = ""
    l_sSendData = ""
    g_Buffer = ""
    
    l_sCommandBuffer = ConvertHexAscii2Hex(Hex(Val(txtDeviceID.Text)))
    
    l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(Val(Mid(cmbFunctioCode.Text, 1, 2))))
    
    l_lTemp = Int(Val(txtStartRegister))
    g_RegisterAddress = l_lTemp
    If l_lTemp > 0 Then l_lTemp = l_lTemp - 1
    

    If l_lTemp > 255 Then
        l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(l_lTemp))
    Else
        l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(0))
        l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(l_lTemp))
    End If

    g_RgisterCount = Val(txtRegisterCount.Text)
    l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(Val(0)))
    l_sCommandBuffer = l_sCommandBuffer & ConvertHexAscii2Hex(Hex(g_RgisterCount))
    
    l_iCRC = MODBUS_CRC_16(l_sCommandBuffer, 6)
    l_sCommandBuffer = l_sCommandBuffer & StrReverse(ConvertHexAscii2Hex(Hex(l_iCRC)))
    
    a_sCommand = l_sCommandBuffer
    
    g_ReceiveIndex = 0
    g_ProgressBarIndex = 0
    g_ProgressBarMaxValue = 5 + (g_RgisterCount * 2)
    ProgressBar1.Value = g_ProgressBarIndex
    ProgressBar1.Max = g_ProgressBarMaxValue
    
    
    ''''MSComm1.Output = l_sCommandBuffer
    WriteToPort l_sCommandBuffer
    
    g_PollesCount = g_PollesCount + 1
    'Text_Polles_Count = g_PollesCount
    StatusBar1.Panels(4).Text = "Polles:  " & Val(g_PollesCount)
    
    Timer_Receive.Interval = Val(TimeOut_SlaveResponse)
    Timer_Receive.Enabled = True
    
    End Sub

Private Sub WriteToPort(sBuffer As String)
    Dim iTemp As Integer
    Dim iVal As Integer
    
    Picture1.Print ">>";
    For iTemp = 1 To Len(sBuffer)
        iVal = Asc(Mid(sBuffer, iTemp, 1))
        Picture1.Print PadLeft(Hex(iVal), 2, "0") & " ";
    Next iTemp
    
    Picture1.Print
    Picture1.Print "<<";
    
    MSComm1.Output = sBuffer
End Sub


Private Function ConverString2Decimal(sData As String) As String
    Dim iPose As Integer
    Dim lTemp As Long
    Dim l_iTemp As String
    
    lTemp = 0
    For iPose = Len(sData) To 1 Step -1
        l_iTemp = Asc(Mid(sData, iPose, 1))
        If CInt(l_iTemp) > 64 Then
            l_iTemp = l_iTemp - 55
        Else
            l_iTemp = Mid(sData, iPose, 1)
        End If
        lTemp = lTemp + CInt(l_iTemp) * (16 ^ (Len(sData) - iPose))
    Next iPose
    
    ConverString2Decimal = lTemp
End Function

Public Function LockParameter()
    Frame_Device_Specifications.Enabled = False
    Frame_Port_Setting.Enabled = False
    Frame_BaudRate.Enabled = False
    Frame_Comm_Mode.Enabled = False
    Time_Config.Enabled = False
    Frame_Funtion_Code.Enabled = False
End Function

Public Function UnLockParameter()
    Frame_Device_Specifications.Enabled = True
    Frame_Port_Setting.Enabled = True
    Frame_BaudRate.Enabled = True
    Frame_Comm_Mode.Enabled = True
    Time_Config.Enabled = True
    Frame_Funtion_Code.Enabled = True
End Function

Public Function DisplayParametersValueHex(a_sReceive() As Byte, a_cTextStart, a_cTextLength As Byte) As String
    Dim l_cTextCount As Byte
    Dim l_sTestDisplay As String
    
    l_sTestDisplay = ""
    
    For l_cTextCount = 0 To a_cTextLength - 1
        l_sTestDisplay = l_sTestDisplay & PadLeft(Hex(a_sReceive(l_cTextCount + a_cTextStart)), 2, "0")
    Next
     l_sTestDisplay = l_sTestDisplay
   
    DisplayParametersValueHex = l_sTestDisplay
End Function

'''''Public Function FillDisplay(l_DisplayArray() As Byte, l_lRegisterAddress As Long, l_iNoOfRegisters As Byte)
'''''    Dim l_iDisplayStartIndex As Integer
'''''    Dim l_lLastRegisterAddress As Long
'''''    Dim l_lFirstRegisterAddress As Long
'''''    Dim l_lLeftIndex As Byte
'''''
'''''    l_iDisplayStartIndex = 3
'''''    l_lFirstRegisterAddress = l_lRegisterAddress
'''''    l_lLastRegisterAddress = l_lFirstRegisterAddress + l_iNoOfRegisters - 1
'''''    l_lLeftIndex = 0
'''''    ResetDisplay
'''''    For l_lRegisterAddress = l_lFirstRegisterAddress To l_lLastRegisterAddress
'''''        Select Case (l_lRegisterAddress Mod 4)
'''''            '-------------------------------------------------------------------------------------------------------------------------
'''''            Case 1
'''''                AddTextField (l_lLeftIndex)
'''''                Text_Display(Text_Display.Count - 1).Text = ConverString2Decimal(DisplayParametersValueHex(l_DisplayArray, l_iDisplayStartIndex, DISP_KWH_LENGTH_BYTE)) / 10000
'''''                Label_Text_Display(Label_Text_Display.Count - 1) = "kWh Ch: " & ((Val(Label_Text_Display.Count) \ 3) + 1)
'''''                l_iDisplayStartIndex = l_iDisplayStartIndex + DISP_KWH_LENGTH_BYTE
'''''                l_lRegisterAddress = l_lRegisterAddress + DISP_KWH_LENGTH_BYTE \ 4
'''''                  '-------------------------------------------------------------------------------------------------------------------------
'''''            Case 2
'''''                AddTextField (l_lLeftIndex)
'''''                Text_Display(Text_Display.Count - 1).Text = "Invalid"
'''''                Label_Text_Display(Label_Text_Display.Count - 1) = "Reg. Length"
'''''                l_iDisplayStartIndex = l_iDisplayStartIndex + DISP_KWH_INVALID_LENGTH_BYTE
'''''            '-------------------------------------------------------------------------------------------------------------------------
'''''            Case 3
'''''                AddTextField (l_lLeftIndex)
'''''                Text_Display(Text_Display.Count - 1).Text = ConverString2Decimal(DisplayParametersValueHex(l_DisplayArray, l_iDisplayStartIndex, DISP_KW_LENGTH_BYTE)) / 100
'''''                Label_Text_Display(Label_Text_Display.Count - 1) = "Power Ch: " & Val(Label_Text_Display.Count) \ 3
'''''                l_iDisplayStartIndex = l_iDisplayStartIndex + DISP_KW_LENGTH_BYTE
'''''            '-------------------------------------------------------------------------------------------------------------------------
'''''            Case 0
'''''                AddTextField (l_lLeftIndex)
'''''                Text_Display(Text_Display.Count - 1).Text = ConverString2Decimal(DisplayParametersValueHex(l_DisplayArray, l_iDisplayStartIndex, DISP_CURRENT_LENGTH_BYTE)) / 100
'''''                Label_Text_Display(Label_Text_Display.Count - 1) = "Current Ch: " & Val(Label_Text_Display.Count) \ 3
'''''                l_iDisplayStartIndex = l_iDisplayStartIndex + DISP_CURRENT_LENGTH_BYTE
'''''            '-------------------------------------------------------------------------------------------------------------------------
'''''            Case Else
'''''                AddTextField (l_lLeftIndex)
'''''                Text_Display(Text_Display.Count - 1).Text = "No Data"
'''''                Label_Text_Display(Label_Text_Display.Count - 1) = "Reg. Length"
'''''                l_iDisplayStartIndex = l_iDisplayStartIndex + DISP_KWH_INVALID_LENGTH_BYTE
'''''        End Select
'''''        l_lLeftIndex = l_lLeftIndex + 1
'''''    Next
'''''End Function
'''''
'''''Public Function ResetDisplay()
'''''    Dim l_cTemp As Byte
'''''    Dim l_cRemoveCount As Byte
'''''    If Text_Display.Count > 1 Then
'''''        l_cRemoveCount = Text_Display.Count - 1
'''''        For l_cTemp = 0 To l_cRemoveCount
'''''            RemoveTextField
'''''        Next
'''''    End If
'''''
'''''End Function

Public Function ResetStatusBar()
    g_PollesCount = 0
    g_ResponseCount = 0
    StatusBar1.Panels(1).Text = "Communication Port:  COM" & Val(g_Port)
    StatusBar1.Panels(2).Text = "Baudrate:  " & Val(g_BaudRate) & "bps"
    StatusBar1.Panels(3).Text = "Request Mode:  " & "Polling"
    StatusBar1.Panels(4).Text = "Polles:  " & Val(g_PollesCount)
    StatusBar1.Panels(5).Text = "Response:  " & Val(g_ResponseCount)
End Function

''''Public Function AddTextField(a_LeftIndex As Integer)
''''    AddTextLabel (a_LeftIndex)
''''    Load Text_Display(Text_Display.Count)
''''    With Text_Display(Text_Display.Count - 1)
''''        .Left = Text_Display(1).Left + ((a_LeftIndex \ DISP_TEXT_BOX_PER_COLUME) * (Text_Display(1).Width + 800))
''''        .Top = Text_Display(1).Top + ((a_LeftIndex Mod DISP_TEXT_BOX_PER_COLUME) * (Text_Display(1).Height + Label_Text_Display(1).Height + 400))
''''        .Visible = True
''''    End With
''''End Function

Public Function RemoveTextField()
    If Text_Display.Count > 1 Then
    Unload Text_Display(Text_Display.Count - 1)
    RemoveTextLabel
    End If
End Function


''''Public Function AddTextLabel(a_LeftIndex As Integer)
''''    Load Label_Text_Display(Label_Text_Display.Count)
''''    With Label_Text_Display(Label_Text_Display.Count - 1)
''''        .Left = Label_Text_Display(1).Left + ((a_LeftIndex \ DISP_TEXT_BOX_PER_COLUME) * (Text_Display(0).Width + 800))
''''        .Top = Label_Text_Display(0).Top + ((a_LeftIndex Mod DISP_TEXT_BOX_PER_COLUME) * (Label_Text_Display(0).Height + Text_Display(0).Height + 400))
''''        .Visible = True
''''    End With
''''End Function

Public Function RemoveTextLabel()
    Unload Label_Text_Display(Label_Text_Display.Count - 1)
End Function






