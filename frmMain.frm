VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4530
   ClientLeft      =   2340
   ClientTop       =   2610
   ClientWidth     =   6030
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picExit 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5520
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   120
      Width           =   285
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   12582912
      MouseIcon       =   "frmMain.frx":0705
      TabCaption(0)   =   "Splash"
      TabPicture(0)   =   "frmMain.frx":0721
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Chat"
      TabPicture(1)   =   "frmMain.frx":073D
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "RTFText"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtMessage"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdSend"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Setup"
      TabPicture(2)   =   "frmMain.frx":0759
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Check1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Check2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "frmMain.frx":0775
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.CheckBox Check2 
         Caption         =   "Save Settings"
         Height          =   255
         Left            =   -70680
         TabIndex        =   34
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Play Sounds"
         Height          =   255
         Left            =   -72000
         TabIndex        =   32
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Default         =   -1  'True
         Height          =   375
         Left            =   4680
         TabIndex        =   27
         ToolTipText     =   "Send"
         Top             =   3120
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Show Border"
         Height          =   855
         Left            =   -72000
         TabIndex        =   20
         Top             =   1680
         Width           =   2655
         Begin VB.OptionButton OptBorderYes 
            Caption         =   "Yes"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            ToolTipText     =   "Show Border"
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton OptBorderNo 
            Caption         =   "No"
            Height          =   255
            Left            =   1440
            TabIndex        =   21
            ToolTipText     =   "No Border"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Choose YES to Move The Screen"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.TextBox txtMessage 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   4335
      End
      Begin VB.Frame Frame6 
         Caption         =   "IP / Connection Information"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   2655
         Begin VB.TextBox txtPort 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   120
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "8205"
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox txtRemoteIP 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txtLocalIP 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "Port Number"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Remote IP Address"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Local IP Address"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Send Popup Message"
         Height          =   735
         Left            =   -74880
         TabIndex        =   11
         Top             =   2880
         Width           =   5535
         Begin VB.TextBox txtPopup 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   3975
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "Send Popup"
            Height          =   375
            Left            =   4200
            TabIndex        =   12
            ToolTipText     =   "Send Popup"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Winsock Commands"
         Height          =   1095
         Left            =   -72000
         TabIndex        =   7
         Top             =   480
         Width           =   2655
         Begin VB.CommandButton cmdListen 
            Caption         =   "Listen"
            Height          =   375
            Left            =   1320
            TabIndex        =   23
            ToolTipText     =   "Listen"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Chat"
            Height          =   375
            Left            =   1320
            TabIndex        =   10
            ToolTipText     =   "Clear InBox"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "Connect"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Connect"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdDisconnect 
            Caption         =   "Disconnect"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Disconnect"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Help"
         Height          =   3015
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   5295
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "fordpref@home.com"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   2520
            Width           =   4935
         End
         Begin VB.Label Label8 
            Caption         =   $"frmMain.frx":0791
            Height          =   495
            Left            =   120
            TabIndex        =   30
            Top             =   1920
            Width           =   4935
         End
         Begin VB.Label Label7 
            Caption         =   $"frmMain.frx":081A
            Height          =   855
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   4935
         End
         Begin VB.Label Label6 
            Caption         =   "If you're having problems using X-Conn, make sure that your Internet or LAN connection is working.  "
            Height          =   495
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Splash Screen"
         Height          =   3135
         Left            =   -74760
         TabIndex        =   3
         Top             =   420
         Width           =   5295
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            Height          =   2535
            Left            =   240
            Picture         =   "frmMain.frx":08F0
            ScaleHeight     =   2475
            ScaleWidth      =   4755
            TabIndex        =   4
            ToolTipText     =   "X-Conn UDP Chat"
            Top             =   360
            Width           =   4815
         End
      End
      Begin RichTextLib.RichTextBox RTFText 
         Height          =   2535
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4471
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":9C04
      End
      Begin VB.Shape Shape1 
         Height          =   135
         Left            =   -75000
         Top             =   3840
         Width           =   735
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Status Bar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "X-Conn - TCP Chat"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Leave these settings alone to play sounds within a VB program
Private Declare Function mciSendString Lib "winmm.dll" Alias _
        "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
        lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
        hwndCallback As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub cmdClear_Click()
    On Error GoTo err ' If there is an error, go to our error trap at the bottom of the subroutine
    mbox = MsgBox("Clear the chat InBox?", vbOKCancel, "Clear InBox?") ' Ask the user if they're sure they want to erase the incoming message box
    If mbox = vbOK Then ' If the user click OK to clear the incoming message box, then...
        RTFText.Text = "" ' Set the incoming message box (RTFText.Text) to "" (nothing)
        txtMessage.SetFocus ' Set the focus back on the outgoing message box (txtMessage) to send another message to the other computer
        Exit Sub ' Exit the subroutine
    End If

err: ' If the user pressed Cancel on the message box above, we end up here, since this produces an error in Visual Basic
    txtMessage.SetFocus ' The user pressed Cancel, so we do nothing but reset the focus back to the outgoing message box
    Exit Sub ' Exit the subroutine
End Sub

Private Sub cmdConnect_Click()
    On Error Resume Next ' If there's an error, resume the next command
    Winsock1.Close ' Close any open ports (just in case)
    Winsock1.Connect txtRemoteIP.Text, txtPort.Text ' Try to connect to the computer IP address specified in the txtRemoteIP text box, on the port specified in the txtPort text box
    lblStatus.Caption = "Connecting to " + txtRemoteIP.Text ' Inform the user we are trying to connect to the specified IP address
    cmdDisconnect.Enabled = True ' Enable the Disconnect button since we may want to disconnect or stop trying to connect
    cmdListen.Enabled = False ' Disable the Listen button and the Connect button since we are already trying to connect
    cmdConnect.Enabled = False ' We are trying to connect, so hide the connect button
    If err Then lblStatus.Caption = err.Description ' If there are any errors, inform the user by showing it on the lblStatus bar
End Sub

Private Sub cmdDisconnect_Click()
    On Error Resume Next ' If there's an error, resume next command
    Winsock1.Close ' We want to disconnect or stop listening for a connection request, so close the connected or listening port
    cmdConnect.Enabled = True ' Enable the Connect button so we can connect to another computer
    cmdListen.Enabled = True ' Enable the Listen button so we can listen for a connection request
    cmdDisconnect.Enabled = False ' We are not connected to anything, so disable the Disconnect button
    lblStatus.Caption = "Disconnected - Not Listening For Request." ' Show the user we are disconnected, and that we are not listening for a connection request
End Sub

Private Sub cmdListen_Click()
    On Error Resume Next ' If there's an error, resume next command
    cmdConnect.Enabled = False ' We are listening for a connection, so disable the Connect button
    cmdListen.Enabled = False ' We are already listening for a connection, so disable the Listen button
    cmdDisconnect.Enabled = True ' Enable the Disconnect button in case you want to stop listening for connection request
    Winsock1.LocalPort = txtPort.Text ' Set the local port to listen on by getting the value from the txtPort text box
    Winsock1.Listen ' Listen for the connection request by the other computer
    lblStatus.Caption = "Listening For Connection Request" ' Inform the user that we are listening for a connection request
End Sub

Private Sub cmdPopup_Click()
    On Error Resume Next ' If there's an error, continue with next command
    If txtPopup.Text = "" Then Exit Sub ' If the txtPopup text box is empty, don't send any data, exit subroutine
    Winsock1.SendData ("|" & txtPopup.Text) ' Send the data with a pipe, |, as first character.  The pipe tells our program that this message is for a popup message box.  See DATA_ARRIVAL subroutine
    txtPopup.Text = "" ' Set the txtPopup box to blank for another popup message
End Sub

Private Sub cmdSend_Click()
    On Error GoTo err ' If there is an error in this subroutine, go to "err" code at bottom
    If txtMessage.Text = "" Then Exit Sub ' If the message is blank, don't send it
    If Winsock1.State = 0 Then ' If we are not connected, do not attempt to send any messages
        txtMessage.Text = "" ' Clear message box
        txtMessage.SetFocus ' and set the focus back to the message box to send another message
        Exit Sub ' Exit subroutine
    End If
    Winsock1.SendData (txtMessage.Text) ' Send our message to the other computer
    RTFText.SelStart = Len(RTFText.Text) ' Set cursor to end of incoming message box. This keeps the last message on the screen
    RTFText.SelColor = &H80FF80 ' Make sure our text is green on the incoming message box
    RTFText.SelText = txtMessage.Text & vbCrLf ' Add current message to the incoming message box (RTFText)
    txtMessage.Text = "" ' Reset the outgoing message box
    txtMessage.SetFocus ' Set the focus back on the outgoing message box to send another message
    Exit Sub ' Exit subroutine
err: ' This is our error trap section for this routine
    lblStatus.Caption = err.Description ' Show the error to the user on the status bar
    txtMessage.Text = "" ' Reset the outgoing message box to blank
    txtMessage.SetFocus ' Set the focus to the outgoing message box to send another message
End Sub

Private Sub Form_Load()
    On Error Resume Next ' If there's an error, continue with next command
    ' Get the settings from the registry if saved, and display them in the appropriate places
    Check1.Value = GetSetting("X-Conn", "Startup", "PlaySounds", 1) ' Play sounds by default
    txtPort.Text = GetSetting("X-Conn", "Startup", "Port", 8205) ' Set the default port to 8205 if nothing is stored in the registry
    txtRemoteIP.Text = GetSetting("X-Conn", "Startup", "Remote", "") ' Clear the txtRemoteIP box if no remote IP is stored in the registry
    Check2.Value = GetSetting("X-Conn", "Startup", "SaveSettings", "True") ' Turn "Save Settings" box on by default, unless specified "False" in the registry
    RTFText.SelColor = &H80FF80 ' Sets foreground color to pale green
    txtMessage.ForeColor = &H80FF80 ' Sets foreground color to pale green
    RTFText.BackColor = &H0& ' Sets background color to black
    txtMessage.BackColor = &H0& ' Sets background color to black
    OptBorderNo.Value = True ' Turns the form border off on load
    txtLocalIP.Text = Winsock1.LocalIP 'Prints the local IP in the txtLocalIP text box
    Winsock1.Close ' Ensures that the winsock port is closed
End Sub

Private Sub OptBorderNo_Click()
    If OptBorderNo.Value = True Then Me.Caption = "" ' Turn off the form border
End Sub

Private Sub OptBorderYes_Click()
    If OptBorderYes.Value = True Then Me.Caption = "X-Conn TCP Chat" ' Turn on the form border (must be on to reposition the form)
End Sub

Private Sub picExit_Click()
    Winsock1.Close ' Close the socket before exiting
    If Check2.Value = 1 Then ' If the "Save Settings" box is checked, save the current settings
            SaveSetting "X-Conn", "Startup", "PlaySounds", Check1.Value ' Saves option to play sounds or not
            SaveSetting "X-Conn", "Startup", "Port", txtPort.Text ' Saves the current port to connect and listen on
            SaveSetting "X-Conn", "Startup", "Remote", txtRemoteIP.Text ' Saves the current computer's remote IP address for quick access next time you run the program
            SaveSetting "X-Conn", "Startup", "SaveSettings", Check2.Value ' Update the registry setting to "Save Settings" for the next program start
    Else ' If the "Save Settings" check box is not checked, do nothing
    End If ' End the sub to save to the registry or not
    Unload Me ' Exits the program
End Sub

Private Sub txtMessage_Change()
    Dim playsound As Long ' Declare the variable to hold the sound to be played if "Play Sounds" box is checked
    If Check1.Value = 1 Then
        playsound = sndPlaySound("xctype.wav", 1) ' If the "Play Sounds" box is checked, play the sound
    End If
End Sub

Private Sub Winsock1_Close()
    lblStatus.Caption = "Connection Has Been Closed." ' Show the user that the connection is closed
    cmdConnect.Enabled = True ' Reset the command buttons.
    cmdListen.Enabled = True  ' Connect and listen need to be enabled
    cmdDisconnect.Enabled = False ' Disable Disconnect since we're not connected or listening for connection
    cmdConnect.SetFocus ' Set the focus back to the cmdConnect button
End Sub

Private Sub Winsock1_Connect()
    On Error Resume Next ' If there's an error, continue with next command
    Dim playsound As Long ' Declare variable to hold the sound to be played if "Play Sounds" box is checked
    lblStatus.Caption = "Connection Has Been Established!" ' Show the user we have a connection
    txtRemoteIP.Text = Winsock1.RemoteHostIP ' Put the remote computer's IP in the remoteIP box
    cmdConnect.Enabled = False ' Disable the Connect and Listen buttons
    cmdListen.Enabled = False  ' We don't need these buttons enabled, and it prevents possible errors
    cmdDisconnect.Enabled = True ' We are connected, so enable the Disconnect button
    If Check1.Value = 1 Then
        playsound = sndPlaySound("xcestab.wav", 1) ' If the "Play Sounds" check box is checked, play the default sound
    End If
    txtMessage.SetFocus ' Set the focus on the box to enter messages to send to the other computer
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next ' Just in case there's an error, continue with next command
    Dim playsound As Long ' Declare variable to hold the sound to be played if "Play Sounds" box is checked
    Winsock1.Close ' Close any open socket (just in case)
    Winsock1.Accept requestID ' Accept the other computer's connection request
    lblStatus.Caption = "Connection Has Been Established!" ' Show the user we have accepted the connection request, and are connected
    txtRemoteIP.Text = Winsock1.RemoteHostIP ' Show the remote computer's IP in the txtRemoteIP text box
    cmdConnect.Enabled = False ' We are connected, so disable the Connect and Listen buttons
    cmdListen.Enabled = False ' This helps to prevent anyone from clicking them and causing errors
    cmdDisconnect.Enabled = True ' Enable the Disconnect button since we're connected
    If Check1.Value = 1 Then
        playsound = sndPlaySound("xcestab.wav", 1) ' If the "Play Sounds" check box is checked, play the default sound
    End If
    txtMessage.SetFocus ' Set the focus on the box to enter messages to send to the other computer
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim playsound As Long ' Declare a variable to hold the sound to play if "PlaySounds" checkbox is checked.
    Dim ndata As String ' Declare a variable to hold the incoming data
    On Error Resume Next ' If there's an error, resume next command
    Winsock1.GetData ndata ' Get the incoming data and store it in variable "ndata"
                ' Check to see if this is a message box
                If Mid(ndata, 1, 1) = "|" Then GoTo boxcode ' Check to see if this is a message box, if it is go to subroutine "boxcode"
    RTFText.SelStart = Len(RTFText.Text) ' Set the cursor to the end of the text box to hold the incoming messages
    RTFText.SelColor = &HC0E0FF ' Set the color of the incoming message to pale orange
    RTFText.SelText = Mid(ndata, InStr(1, ndata, ":") + 1) & vbCrLf ' Insert the text to the last of our current message box (RTFText.text)
    RTFText.SelStart = Len(RTFText.Text) ' Set the cursor to the end of the text box
    RTFText.SelColor = &H80FF80 ' Change the color of the letters back to pale green
    If Check1.Value = 1 Then
        playsound = sndPlaySound("xcmsg.wav", 1) ' If the "Play Sounds" check box is checked, play the default sound
    End If
    Exit Sub ' Exit the subroutine
boxcode: ' If the incoming data's first character was a pipe , |, then the program jumps here
    MsgBox Mid(ndata, 2, Len(ndata) - 1), vbInformation, "Incoming Message" ' Display the incoming data as a message box
    txtMessage.SetFocus ' Put the focus back on the txtmessage box to send another message
    Exit Sub ' Exit the subroutine
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblStatus.Caption = Description ' If there was a winsock error, show the user
    txtMessage.SetFocus ' Set the focus back on the message box to send another message
End Sub
