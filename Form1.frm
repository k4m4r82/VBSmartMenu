VERSION 5.00
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "SmartMenuXP.ocx"
Begin VB.Form Form1 
   Caption         =   "Demo Membuat Menu dengan objek VBSmart Menu XP"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VBSmartXPMenu.SmartMenuXP SmartMenuXP1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BorderStyle     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlign       =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com
' sumber : http://www.visual-basic.com.ar/vbsmart/library/smartmenuxp/smartmenuxp.htm
'***************************************************************************

Private Function getIcon(ByVal iconName As String) As StdPicture
    Set getIcon = LoadPicture(App.Path + "\Icons\" + iconName + ".ico")
End Function

Private Sub addMenuXP()
    With SmartMenuXP1.MenuItems
        .Add 0, "mnuFile", , "&File"
        .Add "mnuFile", "mnuNew", , "&New", getIcon("new")
        .Add "mnuFile", "mnuOpen", , "&Open", getIcon("open")
        .Add "mnuFile", "mnuClose", , "&Close", getIcon("close")
        .Add "mnuFile", , smiSeparator
        .Add "mnuFile", "mnuSave", , "&Save", getIcon("save")
        .Add "mnuFile", "mnuSaveAs", , "Save &As..."
        .Add "mnuFile", , smiSeparator
        .Add "mnuFile", "mnuPrintPreview", , "Print Pre&view", getIcon("preview")
        .Add "mnuFile", "mnuPrint", , "&Print", getIcon("print")
        .Add "mnuFile", , smiSeparator
        
        .Add "mnuFile", "mnuSendTo", , "Sen&d To"
        .Add "mnuSendTo", "mnuMailRecipient", , "&Mail Recipient", getIcon("mail")
        .Add "mnuSendTo", "mnuMailRecipientReview", , "Mail Re&cipient (for Review)"
        .Add "mnuSendTo", "mnuOnlineMeetingParticipant", , "&Online Meeting Participant"
        .Add "mnuSendTo", "mnuFaxRecipient", , "&Fax Recipient...", getIcon("fax")
        .Add "mnuSendTo", , smiSeparator
        .Add "mnuSendTo", "mnuMicrosoftPowerPoint", , "Microsoft &PowerPoint", getIcon("powerpoint")
        
        .Add "mnuFile", , smiSeparator
        .Add "mnuFile", "mnuExit", , "&Exit"
                
        'TODO : DEFINISI MENU YANG LAIN
        
    End With
End Sub

Private Sub Form_Load()
    Call addMenuXP
End Sub

Private Sub SmartMenuXP1_Click(ByVal ID As Long)
    With SmartMenuXP1.MenuItems
        Select Case .Key(ID)
            Case "mnuNew": 'TODO : something here
            Case "mnuOpen": 'TODO : something here
            Case "mnuClose": 'TODO : something here
            Case "mnuSave": 'TODO : something here
            Case "mnuSaveAs": 'TODO : something here
            Case "mnuPrintPreview": 'TODO : something here
            Case "mnuPrint": 'TODO : something here
            Case "mnuMailRecipient": 'TODO : something here
            Case "mnuMailRecipientReview": 'TODO : something here
            Case "mnuOnlineMeetingParticipant": 'TODO : something here
            Case "mnuFaxRecipient": 'TODO : something here
            Case "mnuMicrosoftPowerPoint": 'TODO : something here
            Case "mnuExit": End
        End Select
    End With
End Sub
