VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Converter by Vasilis Ioannidis"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEuro 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00800080&
      Caption         =   "Change conversion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtEuro2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtDrh2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtDrh 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1032
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lbleuro1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      Caption         =   "(#.00)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblDrh2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      Caption         =   "Drachma:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblEuro2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      Caption         =   "Euro:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Currency Converter From Euro to Greek Drachma and From Greek Drachma to Euro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblDrh 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      Caption         =   "Drachma:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblEuro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      Caption         =   "Euro:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        txtEuro.Visible = False
        txtDrh.Visible = False
        lblEuro.Visible = False
        lblDrh.Visible = False
        txtEuro2.Visible = True
        txtDrh2.Visible = True
        lblEuro2.Visible = True
        lblDrh2.Visible = True
        lbleuro1.Visible = False
    Else
        txtEuro2.Visible = False
        txtDrh2.Visible = False
        lblEuro2.Visible = False
        lblDrh2.Visible = False
        txtEuro.Visible = True
        txtDrh.Visible = True
        lblEuro.Visible = True
        lblDrh.Visible = True
        lbleuro1.Visible = True
    End If
End Sub

Private Sub txtEuro_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(",") Then
        MsgBox "Please inmstead of ',' use '.' for the decimal part.", vbInformation, "EuroConverter by Vasilis Ioannidis."
    End If
End Sub

Private Sub txtDrh2_Change()
'340.75 is the rate of 1 Euro in Greek currency
    txtEuro2.Text = Format(Val(txtDrh2.Text) / 340.75, "0.00")
End Sub

Private Sub txtEuro_Change()
    txtDrh.Text = Format(Val(txtEuro.Text) * 340.75, "0")
    txtDrh.Text = Format(txtDrh.Text, "#,##0")
End Sub
