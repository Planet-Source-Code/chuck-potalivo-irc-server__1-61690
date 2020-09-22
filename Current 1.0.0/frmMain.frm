VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "IRC Server"
   ClientHeight    =   6060
   ClientLeft      =   1365
   ClientTop       =   2040
   ClientWidth     =   14955
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   14955
   Begin MSComctlLib.ListView lstConsole 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Time"
         Text            =   "Time"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Text"
         Text            =   "Text"
         Object.Width           =   16757
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu itmExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Admin"
      Begin VB.Menu itmConnections 
         Caption         =   "&Connections ..."
      End
      Begin VB.Menu itmAdmin_Line_1 
         Caption         =   "-"
      End
      Begin VB.Menu itmConfig 
         Caption         =   "Config ..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
  On Error Resume Next
  
  Me.Caption = App.ProductName & " - " & gstrServerVersion
  
  lstConsole.Height = Me.ScaleHeight - (lstConsole.Top * 2)
  lstConsole.Width = Me.ScaleWidth - (lstConsole.Left * 2)

'  Me.Height = (Me.Height - Me.ScaleHeight) + (lstConsole.Top * 2) + lstConsole.Height
'  Me.Width = (Me.Width - Me.ScaleWidth) + (lstConsole.Left * 2) + lstConsole.Width
  
  lstConsole.ColumnHeaders("Time").Width = 2250
  lstConsole.ColumnHeaders("Text").Width = ((1 - (lstConsole.ColumnHeaders("Time").Width / lstConsole.Width)) * lstConsole.Width) - 600
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSettings
End Sub

Private Sub itmConfig_Click()
  Load frmConfig
  
  frmConfig.Show vbModal, Me
End Sub

Private Sub itmConnections_Click()
  frmConnections.Show vbModal, Me
End Sub

Private Sub itmExit_Click()
  Unload Me
End Sub
