VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect StartUp Manager"
   ClientHeight    =   5835
   ClientLeft      =   3525
   ClientTop       =   1245
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniFrame UniFrame1 
      Height          =   1095
      Left            =   120
      Top             =   4680
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1931
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Hu7o71ng da64n"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniLabel UniLabel7 
         Height          =   255
         Left            =   120
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   450
         Caption         =   $"frmMain.frx":169B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin UniControls.UniLabel UniLabel6 
         Height          =   495
         Left            =   120
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   873
         Caption         =   $"frmMain.frx":16A47
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
   End
   Begin UniControls.UniFrame fm 
      Height          =   2655
      Left            =   840
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4683
      MaskColor       =   16711935
      FrameColor      =   -2147483635
      Style           =   0
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniCommonDialog DiaLog1 
         Left            =   5400
         Top             =   360
         _ExtentX        =   714
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniButton cmdCancel 
         Height          =   375
         Left            =   6000
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "frmMain.frx":16B4A
         Style           =   2
         Caption         =   "Hu3y Bo3"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton cmdAdd 
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "frmMain.frx":170E4
         Style           =   2
         Caption         =   "The6m"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniTextBox txtKeyName 
         Height          =   270
         Left            =   2040
         TabIndex        =   17
         Top             =   1440
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
      End
      Begin UniControls.UniLabel UniLabel5 
         Height          =   255
         Left            =   240
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Te6n Kho1a:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniComboBox cb 
         Height          =   330
         Left            =   2040
         TabIndex        =   16
         Top             =   1080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ExtendedUI      =   0   'False
         DropDownWidth   =   0
         AutoCompleteListItemsOnly=   -1  'True
         AutoCompleteItemsAreSorted=   -1  'True
      End
      Begin UniControls.UniLabel UniLabel4 
         Height          =   255
         Left            =   240
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Vi5 Tri1 Khoa1:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniButton cmdBrowser 
         Height          =   255
         Left            =   6960
         TabIndex        =   15
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Icon            =   "frmMain.frx":1767E
         Style           =   2
         Caption         =   "..."
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniTextBox txtPath 
         Height          =   270
         Left            =   2040
         TabIndex        =   14
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
      End
      Begin UniControls.UniLabel UniLabel3 
         Height          =   255
         Left            =   240
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Chu7o7ng tri2nh the6m:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel2 
         Height          =   255
         Left            =   2280
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "The6m va2o Start Up"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
   End
   Begin UniControls.UniButton cmdCreate 
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmMain.frx":1769A
      Style           =   2
      Caption         =   "The6m va2o"
      IconAlign       =   3
      iNonThemeStyle  =   2
      BackColor       =   -2147483643
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton cmdRefer 
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmMain.frx":17C34
      Style           =   2
      Caption         =   "Refresh"
      IconAlign       =   3
      iNonThemeStyle  =   2
      BackColor       =   -2147483643
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton cmdJump 
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmMain.frx":181CE
      Style           =   2
      Caption         =   "D9i d9e61n"
      IconAlign       =   3
      iNonThemeStyle  =   2
      BackColor       =   -2147483643
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton cmdDel 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   4200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmMain.frx":18768
      Style           =   2
      Caption         =   "Xo1a"
      IconAlign       =   3
      iNonThemeStyle  =   2
      BackColor       =   -2147483643
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton cmdPhucHoi 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmMain.frx":18D02
      Style           =   2
      Caption         =   "Phu5c Ho62i"
      IconAlign       =   3
      iNonThemeStyle  =   2
      BackColor       =   -2147483643
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin UniControls.UniListView LV 
      Height          =   3255
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      BorderStyle     =   2
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   5
      Left            =   360
      Picture         =   "frmMain.frx":1929C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   4
      Left            =   360
      Picture         =   "frmMain.frx":19826
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   3
      Left            =   720
      Picture         =   "frmMain.frx":19DB0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   2
      Left            =   480
      Picture         =   "frmMain.frx":1A33A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   1
      Left            =   360
      Picture         =   "frmMain.frx":1A8C4
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":1AE4E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin UniControls.UniTreeView Tree1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5741
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Qua3n ly1 kho73i d9o65ng"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Sub cmdAdd_Click()
If txtPath.Text = "" Or cb.Text = "" Or txtKeyName.Text = "" Then
    UniMsgBox "Chu7a nha65p d9u3 tho6ng tin!", vbOKOnly + vbCritical
Else
    Dim HY
    If GetKeyGoc(cb.Text) = "HKEY_CURRENT_USER" Then
        HY = &H80000001
    ElseIf GetKeyGoc(cb.Text) = "HKEY_LOCAL_MACHINE" Then
        HY = &H80000001
    End If
    SaveString HY, GetKeyPath(cb.Text), txtKeyName.Text, txtPath.Text
    UniMsgBox "D9a4 the6m xong!", vbOKOnly + vbInformation, "OK"
    fm.Visible = False
    cmdRefer_Click
End If
End Sub

Private Sub cmdBrowser_Click()
On Error GoTo ThOaTdIcHoLe

DiaLog1.FileName = ""
DiaLog1.ShowOpen
txtPath.Text = DiaLog1.FileName



ThOaTdIcHoLe:
End Sub

Private Sub cmdCancel_Click()
fm.Visible = False
End Sub

Private Sub cmdCreate_Click()
fm.Visible = True
End Sub

Private Sub cmdDel_Click()
Dim i
Dim uh As Boolean
uh = False
For i = 1 To LV.ListItems.Count
    If LV.ListItems(i).Selected = True Then uh = True
Next i
If uh = True Then
    If UniMsgBox("Ba5n co1 muo61n xo1a chu7o7ng tri2nh [" & LV.SelectedItem.Text & "] ra kho3i Start Up ngay ba6y gio72 kho6ng?", vbYesNo + vbInformation, "Xo1a") = vbYes Then
    
        If LV.SelectedItem.SubItems(2).Caption <> "---" Then
        
            Dim YH
            Dim AH
                YH = GetKeyGoc(LV.SelectedItem.SubItems(2).Caption)
                If YH = "HKEY_CURRENT_USER" Then
                    AH = &H80000001
                ElseIf YH = "HKEY_LOCAL_MACHINE" Then
                    AH = &H80000002
                End If
                
                DeleteValue AH, GetKeyPath(LV.SelectedItem.SubItems(2).Caption), LV.SelectedItem.Text
                

        Else
        
            SetAttr LV.SelectedItem.SubItems(1).Caption, vbNormal
            DeleteFile LV.SelectedItem.SubItems(1).Caption
            
        End If '"---"
        
        UniMsgBox "D9a4 xo1a xong!", vbOKOnly + vbInformation, "OK"
                
        cmdRefer_Click
    End If ' unimsgbo
Else
    UniMsgBox "Kho6ng co1 kho1a na2o d9ang d9u7o75c cho5n", vbOKOnly + vbCritical
End If 'uh
End Sub
Private Function GetKeyGoc(sKey)
GetKeyGoc = Mid(sKey, 1, InStr(1, sKey, "\") - 1)
End Function
Private Function GetKeyPath(sKey)
GetKeyPath = Mid(sKey, InStr(1, sKey, "\") + 1, Len(sKey) - InStr(1, sKey, "\"))
End Function

Private Sub cmdPhucHoi_Click()

Dim i
Dim uh As Boolean
uh = False
For i = 1 To LV.ListItems.Count
    If LV.ListItems(i).Selected = True Then uh = True
Next i
If uh = True Then
    If UniMsgBox("Ba5n co1 muo61n phu5c ho62i kho1a [" & LV.SelectedItem.Text & "] ngay ba6y gio72 kho6ng?", vbYesNo + vbInformation, "Phu5c Ho62i") = vbYes Then
        If LV.SelectedItem.Text = ToUnicode("Shell [He65 Tho61ng]") Then
            SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe"
        ElseIf LV.SelectedItem.Text = ToUnicode("Userinit [He65 Tho61ng]") Then
            SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe,"
        End If
        
        UniMsgBox "D9a4 phu5c ho62i xong!", vbOKOnly + vbInformation, "OK"
    End If
Else
    UniMsgBox "Kho6ng co1 kho1a na2o d9ang d9u7o75c cho5n", vbOKOnly + vbCritical
End If
End Sub

Private Sub cmdRefer_Click()
LV.ListItems.Clear
GetSystemKey
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
GetFolderStartUp 1
GetFolderStartUp 2
End Sub

Private Sub Form_Load()




With cb
    .AddItem "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    .AddItem "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    .AddItem "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    .AddItem "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    .AddItem "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
    .ListIndex = 0
End With

fm.Visible = False



    File1.System = True
    File1.Hidden = True
    File1.ReadOnly = True
    File1.Archive = True
    File1.Pattern = "*.exe;*.bat;*.cmd;*.com;*.pif;*.dll;*.lnk;*.url"
    
    
With Tree1
    .Initialize
    .InitializeImageList

    .AddIcon Ico(0).Picture
    .AddIcon Ico(1).Picture
    .AddIcon Ico(2).Picture
    .AddIcon Ico(3).Picture
    .AddIcon Ico(4).Picture
    .AddIcon Ico(5).Picture
    
    
    .AddNode , , "a", "Ta61t ca3 ca1c tri2nh kho73i d9o65ng", 0, 0
        .AddNode "a", , "aReg", "Tu72 Registry", 1, 1
            .AddNode "aReg", , "aRegAdmin", "Ta61t ca3 ngu7o72i du2ng", 2, 2
                .AddNode "aRegAdmin", , "aRegAdmin0", "Cha5y", 3, 3
                .AddNode "aRegAdmin", , "aRegAdmin1", "Cha5y 1 La62n", 3, 3
            .AddNode "aReg", , "aRegUser", Environ$("USERNAME"), 2, 2
                .AddNode "aRegUser", , "aRegUser0", "Cha5y", 3, 3
                .AddNode "aRegUser", , "aRegUser1", "Cha5y 1 La62n", 3, 3
        .AddNode "a", , "aFol", "Tu72 Thu7 Mu5c Kho73i D9o65ng", 4, 4
            .AddNode "aFol", , "aFolAdmin", "Ta61t ca3 ngu7o72i du2ng", 2, 2
            .AddNode "aFol", , "aFolUser", Environ$("USERNAME"), 2, 2
        .AddNode "a", , "aSys", "Kho1a He65 Tho61ng", 5, 5
        
        
    .Expand .GetKeyNode(a), True
End With


With LV
    .View = eViewDetails
    .FullRowSelect = True
    .GridLines = True
    .AutoUnicode = False
    .CheckBoxes = False
    .MultiSelect = False
    
    .Columns.Add , , ToUnicode("Te6n Chu7o7ng Tri2nh"), , 2000
    .Columns.Add , , ToUnicode("D9i5a Chi3"), , 4000
    .Columns.Add , , ToUnicode("D9u7o72ng Da64n Key"), , 4000

End With


GetSystemKey
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
GetFolderStartUp 1
GetFolderStartUp 2

End Sub

Private Sub LV_ItemClick(Item As UniControls.cListItem)
If Right(Item.Text, Len(ToUnicode("[He65 Tho61ng]"))) = ToUnicode("[He65 Tho61ng]") Then
    cmdPhucHoi.Enabled = True
    cmdDel.Enabled = False
Else
    cmdPhucHoi.Enabled = False
    cmdDel.Enabled = True
End If
End Sub

Private Sub Tree1_NodeClick(ByVal hNode As Long)
LV.ListItems.Clear
Dim sKey
sKey = Tree1.GetNodeKey(hNode)
If sKey = "aReg" Then
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
ElseIf sKey = "aRegAdmin" Then
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
ElseIf sKey = "aRegAdmin0" Then
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
ElseIf sKey = "aRegAdmin1" Then
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
ElseIf sKey = "aRegUser" Then
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
ElseIf sKey = "aRegUser0" Then
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
ElseIf sKey = "aRegUser1" Then
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
ElseIf sKey = "aFol" Then
    GetFolderStartUp 1
    GetFolderStartUp 2
ElseIf sKey = "aFolAdmin" Then
    GetFolderStartUp 1
ElseIf sKey = "aFolUser" Then
    GetFolderStartUp 2
ElseIf sKey = "aSys" Then
    GetSystemKey
ElseIf sKey = "a" Then
    GetSystemKey
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
    GetFolderStartUp 1
    GetFolderStartUp 2
End If
End Sub

