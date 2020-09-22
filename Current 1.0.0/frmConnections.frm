VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConnections 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connections"
   ClientHeight    =   4605
   ClientLeft      =   2565
   ClientTop       =   3090
   ClientWidth     =   10980
   Icon            =   "frmConnections.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   9660
      Top             =   720
   End
   Begin MSComctlLib.ListView lstConnections 
      Height          =   4275
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   7541
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ident"
         Object.Width           =   7497
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nickname"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Remote IP"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Local IP"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Ping"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   9660
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmConnections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadList()
  Static bolRunning As Boolean
  
  If bolRunning Then
    Exit Sub
  End If
  bolRunning = True

  Dim itmConnect As ListItem
  
  For Each itmConnect In lstConnections.ListItems
    itmConnect.Checked = False
  Next
  
  Dim conListItem As clsConnect
  
  Dim lngIndex As Long
  lngIndex = 1
  
  On Error Resume Next
  
  Do Until lngIndex > colConnects.Count
    Set conListItem = colConnects(lngIndex)
    
    Set itmConnect = Nothing
    Set itmConnect = lstConnections.ListItems(conListItem.GUID)
    
    If itmConnect Is Nothing Then
      Set itmConnect = lstConnections.ListItems.Add(, conListItem.GUID)
    End If
    
    itmConnect.Text = conListItem.FullIdent
    itmConnect.SubItems(1) = conListItem.Nick
    itmConnect.SubItems(2) = conListItem.RemoteIP
    itmConnect.SubItems(3) = conListItem.LocalIP
    itmConnect.SubItems(4) = conListItem.Ping & " ms"
    
    itmConnect.Checked = True
    
    lngIndex = lngIndex + 1
  Loop
  
  lngIndex = 1
  
  Do Until lngIndex > lstConnections.ListItems.Count
    Set itmConnect = lstConnections.ListItems(lngIndex)
    
    If Not itmConnect.Checked Then
      lstConnections.ListItems.Remove itmConnect.Key
    Else
      lngIndex = lngIndex + 1
    End If
  Loop
  
  Me.Caption = "Connections - " & colConnects.Count
  
  bolRunning = False
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  LoadList

  Me.Width = (Me.Width - Me.ScaleWidth) + (lstConnections.Left * 3) + lstConnections.Width + cmdClose.Width
  Me.Height = (Me.Height - Me.ScaleHeight) + (lstConnections.Top * 2) + lstConnections.Height
  
  cmdClose.Left = (lstConnections.Left * 2) + lstConnections.Width
  cmdClose.Top = lstConnections.Top
End Sub

Private Sub tmrUpdate_Timer()
  LoadList
End Sub
