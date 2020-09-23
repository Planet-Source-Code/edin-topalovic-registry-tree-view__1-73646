VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegistryView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RegView"
   ClientHeight    =   5160
   ClientLeft      =   3195
   ClientTop       =   3540
   ClientWidth     =   9735
   Icon            =   "frmRegistryExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9735
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   4890
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11986
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   2
            TextSave        =   "15:13"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   2
            TextSave        =   "22.12.2010"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   6480
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   6000
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5160
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilSmall 
      Left            =   600
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistryExample.frx":000C
            Key             =   "fldrClosed"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilMain 
      Left            =   0
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistryExample.frx":05A6
            Key             =   "fldrClosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistryExample.frx":0B40
            Key             =   "fldrOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistryExample.frx":10DA
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistryExample.frx":1674
            Key             =   "explorer"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   8705
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ilMain"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4935
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "il1"
      SmallIcons      =   "il2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "ime"
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Sadržaj"
         Text            =   "Data"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmRegistryView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
MsgBox "Created by Edin Topalovic" & vbCrLf & "Email: topalovic.e@gmail.com", vbInformation
End Sub


Private Sub Form_Load()
  On Error Resume Next
  Dim Message As String
  Dim Counter As Integer
tv.Nodes.Add , , "MyComputer", "My Computer", "drive"
tv.Nodes(1).Expanded = True
  tv.Nodes.Add "MyComputer", tvwChild, "HKEY_CLASSES_ROOT", "HKEY_CLASSES_ROOT", 1, 2
  tv.Nodes.Add "MyComputer", tvwChild, "HKEY_CURRENT_USER", "HKEY_CURRENT_USER", 1, 2
  tv.Nodes.Add "MyComputer", tvwChild, "HKEY_LOCAL_MACHINE", "HKEY_LOCAL_MACHINE", 1, 2
  tv.Nodes.Add "MyComputer", tvwChild, "HKEY_USERS", "HKEY_USERS", 1, 2
  tv.Nodes.Add "MyComputer", tvwChild, "HKEY_CURRENT_CONFIG", "HKEY_CURRENT_CONFIG", 1, 2
  Message = mR.EnumerateRegKeys1(HKEY_CLASSES_ROOT, "")
  tv.Nodes.Add "HKEY_CLASSES_ROOT", tvwChild, "HKEY_CLASSES_ROOT\" & Message, Message, 1, 2
  Message = mR.EnumerateRegKeys1(HKEY_CURRENT_USER, "")
  tv.Nodes.Add "HKEY_CURRENT_USER", tvwChild, "HKEY_CURRENT_USER\" & Message, Message, 1, 2
  Message = mR.EnumerateRegKeys1(HKEY_LOCAL_MACHINE, "")
  tv.Nodes.Add "HKEY_LOCAL_MACHINE", tvwChild, "HKEY_LOCAL_MACHINE\" & Message, Message, 1, 2
  Message = mR.EnumerateRegKeys1(HKEY_USERS, "")
  tv.Nodes.Add "HKEY_USERS", tvwChild, "HKEY_USERS\" & Message, Message, 1, 2
  Message = mR.EnumerateRegKeys1(HKEY_CURRENT_CONFIG, "")
  tv.Nodes.Add "HKEY_CURRENT_CONFIG", tvwChild, "HKEY_CURRENT_CONFIG\" & Message, Message, 1, 2
  cnt = 1
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
Dim pth As String
Dim hk As KeyRoot
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    If cnt1 = UBound(NewMessage1) + 1 Then
        Timer1.Enabled = False
    End If
        If knm1 = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm1 = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm1 = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm1 = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm1 = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
    pth = knm1 & "\" & NewMessage1(cnt1)
    tv.Nodes.Add knm1, tvwChild, pth, NewMessage1(cnt1), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm1 & "\" & NewMessage1(cnt1)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    pb.value = cnt1
    cnt1 = cnt1 + 1
    
End Sub

Private Sub Timer2_Timer()
If pb.value = pb.Max Then
    Me.Height = Me.Height - pb.Height
    Timer2.Enabled = False
End If
End Sub

Private Sub tv_Expand(ByVal Node As MSComctlLib.Node)
Dim pth As String
Dim hk As KeyRoot
Dim ll As String
Dim km As String
With Node
For i = 0 To List1.ListCount - 1
    If List1.List(i) = .key Then Exit Sub
Next i
If .Text = "My Computer" Then

ElseIf .Text = "HKEY_CLASSES_ROOT" Then
    'tv.Enabled = False
    'lv.Enabled = False
    cnt1 = 0
    tv.Nodes.Remove (Node.Child.Index)
    Screen.MousePointer = vbHourglass
    Message = mR.EnumerateRegKeys(HKEY_CLASSES_ROOT, "")
    List1.AddItem "HKEY_CLASSES_ROOT"
    Screen.MousePointer = 0
    NewMessage1 = Split(Message, Chr(0))
    knm1 = "HKEY_CLASSES_ROOT"
    pb.Min = 0
    pb.Max = UBound(NewMessage1)
    pb.value = pb.Min
    Me.Height = Me.Height + pb.Height
    Timer1.Enabled = True
ElseIf .Text = "HKEY_CURRENT_USER" Then
    cnt = 0
    tv.Nodes.Remove (Node.Child.Index)
    Screen.MousePointer = vbHourglass
    Message = mR.EnumerateRegKeys(HKEY_CURRENT_USER, "")
    List1.AddItem "HKEY_CURRENT_USER"
    NewMessage = Split(Message, Chr(0))
    For i = LBound(NewMessage) To UBound(NewMessage)
        hk = HKEY_CURRENT_USER
        knm = "HKEY_CURRENT_USER"
    pth = knm & "\" & NewMessage(i)
    tv.Nodes.Add knm, tvwChild, pth, NewMessage(i), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm & "\" & NewMessage(i)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    Next i
    Screen.MousePointer = 0
ElseIf .Text = "HKEY_LOCAL_MACHINE" Then
    cnt = 0
    tv.Nodes.Remove (Node.Child.Index)
    Screen.MousePointer = vbHourglass
    Message = mR.EnumerateRegKeys(HKEY_LOCAL_MACHINE, "")
    List1.AddItem "HKEY_LOCAL_MACHINE"
    NewMessage = Split(Message, Chr(0))
    For i = LBound(NewMessage) To UBound(NewMessage)
        hk = HKEY_LOCAL_MACHINE
        knm = "HKEY_LOCAL_MACHINE"
    pth = knm & "\" & NewMessage(i)
    tv.Nodes.Add knm, tvwChild, pth, NewMessage(i), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm & "\" & NewMessage(i)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    Next i
    Screen.MousePointer = 0
ElseIf .Text = "HKEY_USERS" Then
    cnt = 0
    tv.Nodes.Remove (Node.Child.Index)
    Screen.MousePointer = vbHourglass
    Message = mR.EnumerateRegKeys(HKEY_USERS, "")
    List1.AddItem "HKEY_USERS"
    NewMessage = Split(Message, Chr(0))
    For i = LBound(NewMessage) To UBound(NewMessage)
        hk = HKEY_USERS
        knm = "HKEY_USERS"
    pth = knm & "\" & NewMessage(i)
    tv.Nodes.Add knm, tvwChild, pth, NewMessage(i), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm & "\" & NewMessage(i)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    Next i
    Screen.MousePointer = 0
ElseIf .Text = "HKEY_CURRENT_CONFIG" Then
    cnt = 0
    tv.Nodes.Remove (Node.Child.Index)
    Screen.MousePointer = vbHourglass
    Message = mR.EnumerateRegKeys(HKEY_CURRENT_CONFIG, "")
    List1.AddItem "HKEY_CURRENT_CONFIG"
    NewMessage = Split(Message, Chr(0))
    For i = LBound(NewMessage) To UBound(NewMessage)
        hk = HKEY_CURRENT_CONFIG
        knm = "HKEY_CURRENT_CONFIG"
    pth = knm & "\" & NewMessage(i)
    tv.Nodes.Add knm, tvwChild, pth, NewMessage(i), 1, 2
    pth = Right(pth, Len(pth) - InStr(1, pth, "\"))
    ms = mR.EnumerateRegKeys1(hk, pth)
    pth = knm & "\" & NewMessage(i)
    If ms <> "" Then
        tv.Nodes.Add pth, tvwChild, pth & "\" & ms, ms, 1, 2
    End If
    Next i
    Screen.MousePointer = 0
Else
    cnt = 0
    Screen.MousePointer = vbHourglass
    ll = Right(.FullPath, Len(.FullPath) - InStr(1, .FullPath, "\"))
    knm = Left(ll, InStr(1, ll, "\") - 1)
        If knm = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf knm = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf knm = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf knm = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf knm = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
        pth = Right(ll, Len(ll) - Len(knm) - 1)
    Message = mR.EnumerateRegKeys(hk, pth)
    List1.AddItem ll
    NewMessage = Split(Message, Chr(0))
    tv.Nodes.Remove (Node.Child.Index)
    For i = LBound(NewMessage) To UBound(NewMessage)
    tv.Nodes.Add knm & "\" & pth, tvwChild, knm & "\" & pth & "\" & NewMessage(i), NewMessage(i), 1, 2
    km = pth & "\" & NewMessage(i)
    ms = mR.EnumerateRegKeys1(hk, km)
    If ms <> "" Then
        tv.Nodes.Add knm & "\" & km, tvwChild, knm & "\" & km & "\" & ms, ms, 1, 2
    End If
    Next i
    Screen.MousePointer = 0
End If
End With
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Dim hk As KeyRoot
Dim pt As String
lv.ListItems.Clear
    Screen.MousePointer = vbHourglass
    sel = tv.SelectedItem.key
    If InStr(sel, "\") = 0 Then
        If sel = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf sel = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf sel = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf sel = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf sel = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
        b = mR.EnumerateRegKeyValues(hk, "")
    Else
        a = sel
        sel = Left(sel, InStr(sel, "\") - 1)
        pt = Right(a, Len(a) - Len(sel) - 1)
        If sel = "HKEY_CLASSES_ROOT" Then
            hk = HKEY_CLASSES_ROOT
        ElseIf sel = "HKEY_CURRENT_CONFIG" Then
            hk = HKEY_CURRENT_CONFIG
        ElseIf sel = "HKEY_CURRENT_USER" Then
            hk = HKEY_CURRENT_USER
        ElseIf sel = "HKEY_LOCAL_MACHINE" Then
            hk = HKEY_LOCAL_MACHINE
        ElseIf sel = "HKEY_USERS" Then
            hk = HKEY_USERS
        End If
        b = mR.EnumerateRegKeyValues(hk, pt)
    End If
    mss = Split(b, Chr(0))
    lv.ListItems.Add , , "(Default)"
    lv.ListItems(lv.ListItems.Count).SubItems(1) = "(value not set)"
    For i = LBound(mss) To UBound(mss)
        k = Split(mss(i), Chr(1))
        If Trim(k(0)) = "" And k(1) <> "" Then
            If i = 0 Then
            lv.ListItems(1).SubItems(1) = k(1)
            GoTo out_:
            End If
        End If
        lv.ListItems.Add , , k(0)
        lv.ListItems(lv.ListItems.Count).SubItems(1) = k(1)
out_:
    Next i
    Screen.MousePointer = 0
End Sub
