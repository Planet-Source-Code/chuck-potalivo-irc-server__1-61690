VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IRC Server Configuration"
   ClientHeight    =   3990
   ClientLeft      =   1545
   ClientTop       =   1725
   ClientWidth     =   8145
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraBlank 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   1980
      TabIndex        =   3
      Top             =   60
      Width           =   4755
   End
   Begin MSComctlLib.TreeView treeConfig 
      Height          =   3855
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   6800
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      FullRowSelect   =   -1  'True
      Scroll          =   0   'False
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  SaveSettings
  LoadSettings

  Unload Me
End Sub

Private Sub Form_Load()
  Dim nodNew              As Node
  
  Set nodNew = treeConfig.Nodes.Add(, , "IRC", "IRC Server")
  nodNew.Expanded = True
  
    Set nodNew = treeConfig.Nodes.Add("IRC", tvwChild, "IRC_Network", "Network")
    Set nodNew = treeConfig.Nodes.Add("IRC", tvwChild, "IRC_Client", "Client")
  
  Set nodNew = treeConfig.Nodes.Add(, , "Telnet", "Telnet Server")
  nodNew.Expanded = True
End Sub
