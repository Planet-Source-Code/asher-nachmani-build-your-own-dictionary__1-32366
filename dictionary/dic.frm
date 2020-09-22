VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form dic 
   Caption         =   " "
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "add a new word to dictionary"
      Height          =   945
      Left            =   60
      TabIndex        =   7
      Top             =   1170
      Width           =   6375
      Begin VB.CommandButton Command1 
         Caption         =   "add"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   420
         Width           =   945
      End
      Begin VB.TextBox txtp 
         Height          =   345
         Left            =   1230
         TabIndex        =   11
         Top             =   450
         Width           =   1875
      End
      Begin VB.TextBox txtNewWord 
         Height          =   345
         Left            =   3420
         TabIndex        =   8
         Top             =   450
         Width           =   1875
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   " explanation to word"
         Height          =   255
         Left            =   2100
         TabIndex        =   10
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "word"
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         Top             =   180
         Width           =   585
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete"
      Height          =   255
      Left            =   4770
      TabIndex        =   5
      Top             =   2130
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "translate"
      Height          =   375
      Left            =   150
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   2085
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "dic.frx":0000
      Left            =   3600
      List            =   "dic.frx":001F
      TabIndex        =   0
      Top             =   390
      Width           =   1935
   End
   Begin MSComctlLib.ListView LV 
      Height          =   1515
      Left            =   30
      TabIndex        =   6
      Top             =   2400
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   2672
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageListSapak"
      SmallIcons      =   "ImageListSapak"
      ColHdrIcons     =   "ImageListSapak"
      ForeColor       =   14573857
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "enter a word to  translate"
      Height          =   255
      Left            =   1500
      TabIndex        =   4
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "choose  dictionary"
      Height          =   315
      Left            =   3930
      TabIndex        =   1
      Top             =   90
      Width           =   1485
   End
End
Attribute VB_Name = "dic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
        Lang = Combo1
End Sub
Private Sub Command1_Click()
If Combo1 = "" Then Lang = ""
If Lang = "" Then
  MsgBox "not choose a dictionary "
  Exit Sub
End If
If ChekDic(Lang) = True Then
Else
    If MsgBox("this dictionary not find do you want to create it", vbRetryCancel) = vbCancel Then
      Exit Sub
    End If
End If

If txtNewWord.Text = "" Then
    If MsgBox("you dont put a word", vbRetryCancel) = vbCancel Then
       Exit Sub
    End If
    Exit Sub
End If
 If txtp.Text = "" Then
    If MsgBox("you dont put an explanation to the word", vbRetryCancel) = vbCancel Then
       
       Exit Sub
    End If
    Exit Sub
End If
    NewTextWord = Trim(txtNewWord.Text)
    Translate = Trim(txtp.Text)
    AddNewWord Lang, NewTextWord, Translate

    TransletWord Lang, txtNewWord.Text
     If ReadNewTextWord(0) <> "" Then
       FillParitim LV, txtNewWord.Text
     Else
       LV.ListItems.Clear
       MsgBox "the translate to this word not found"
       Exit Sub
     End If




End Sub

Private Sub Command2_Click()
If Combo1 = "" Then Lang = ""
 If Lang = "" Then
  MsgBox "not choose a dictionary"
  Exit Sub
End If
If ChekDic(Lang) = True Then
Else
  MsgBox "this dictionary not find"
  Exit Sub
End If
  If Text1.Text = "" Then Exit Sub
  TransletWord Lang, Text1.Text
     If ReadNewTextWord(0) <> "" Then
       FillParitim LV, Text1.Text
     Else
       LV.ListItems.Clear
       MsgBox "the translate to this word not found"
       Exit Sub
     End If
End Sub

Private Sub Command3_Click()
 If Combo1 = "" Then Lang = ""
 If Lang = "" Then
  MsgBox "not choose a dictionary"
  Exit Sub
End If
   If Del = "" Or deltranslet = "" Then
            MsgBox "you shoold pick from the list"
            Exit Sub
   End If
DelWord Lang, Del, deltranslet
  
  TransletWord Lang, Del
     If ReadNewTextWord(0) <> "" Then
       FillParitim LV, Del
     Else
       LV.ListItems.Clear
       Exit Sub
     End If

End Sub


Private Sub Form_Activate()
 Del = ""
deltranslet = ""

End Sub

Private Sub Form_Load()
 Del = ""
deltranslet = ""

  FillListColumnHeader LV
End Sub

Private Sub LV_Click()

On Error Resume Next

        If LV.SelectedItem.Selected <> False Then
            Del = LV.SelectedItem.SubItems(1)
            deltranslet = LV.SelectedItem.SubItems(2)
        End If
   

End Sub
