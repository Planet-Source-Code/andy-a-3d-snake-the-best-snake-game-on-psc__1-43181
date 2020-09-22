VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "3D Snake"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF8080&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin VB.PictureBox picLevel 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1830
         Left            =   4148
         ScaleHeight     =   1830
         ScaleWidth      =   3705
         TabIndex        =   35
         Top             =   6240
         Width           =   3705
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   255
         Left            =   5115
         Max             =   10
         TabIndex        =   32
         Top             =   5880
         Width           =   3495
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   5085
         Max             =   12
         Min             =   4
         TabIndex        =   24
         Top             =   5160
         Value           =   10
         Width           =   3495
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   5040
         Max             =   12
         TabIndex        =   20
         Top             =   4320
         Value           =   12
         Width           =   3495
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   5040
         Max             =   12
         TabIndex        =   16
         Top             =   3480
         Value           =   7
         Width           =   3495
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   5040
         Max             =   20
         Min             =   1
         TabIndex        =   12
         Top             =   2640
         Value           =   14
         Width           =   3495
      End
      Begin Project1.Button Button1 
         Height          =   495
         Left            =   9960
         TabIndex        =   31
         Top             =   8280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OK"
         Shape           =   0
         FillColor       =   12582912
         FillColorMouseOver=   12582912
         FillColorMouseDown=   16711680
         ForeColorInvert =   0   'False
         FontChangeMouseDown=   0
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start at Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2820
         TabIndex        =   34
         Top             =   5760
         Width           =   2010
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8670
         TabIndex        =   33
         Top             =   5835
         Width           =   180
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8640
         TabIndex        =   30
         Top             =   5115
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8640
         TabIndex        =   29
         Top             =   4275
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8640
         TabIndex        =   28
         Top             =   3435
         Width           =   180
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8640
         TabIndex        =   27
         Top             =   2595
         Width           =   360
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5115
         TabIndex        =   26
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8325
         TabIndex        =   25
         Top             =   5520
         Width           =   270
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fruits per Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2610
         TabIndex        =   23
         Top             =   5040
         Width           =   2370
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4905
         TabIndex        =   22
         Top             =   4680
         Width           =   480
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loud"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8085
         TabIndex        =   21
         Top             =   4680
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound Volume:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2595
         TabIndex        =   19
         Top             =   4200
         Width           =   2310
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4965
         TabIndex        =   18
         Top             =   3840
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loud"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8145
         TabIndex        =   17
         Top             =   3840
         Width           =   600
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music Volume:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2700
         TabIndex        =   15
         Top             =   3360
         Width           =   2220
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8145
         TabIndex        =   14
         Top             =   3000
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4875
         TabIndex        =   13
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Snake Speed:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2760
         TabIndex        =   11
         Top             =   2520
         Width           =   2100
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   555
         Left            =   5100
         TabIndex        =   10
         Top             =   1920
         Width           =   1800
      End
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF8080&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin Project1.Button Button3 
         Height          =   495
         Left            =   9960
         TabIndex        =   42
         Top             =   8280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OK"
         Shape           =   0
         FillColor       =   12582912
         FillColorMouseOver=   12582912
         FillColorMouseDown=   16711680
         ForeColorInvert =   0   'False
         FontChangeMouseDown=   0
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Up Arrow: Turn Right"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6330
         TabIndex        =   56
         Top             =   6240
         Width           =   2880
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Down Arrow: Turn Left"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6195
         TabIndex        =   55
         Top             =   6600
         Width           =   3120
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P: Pause/Resume"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6480
         TabIndex        =   54
         Top             =   6960
         Width           =   2550
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Esc.: Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   7005
         TabIndex        =   53
         Top             =   7320
         Width           =   1470
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":2764F
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1200
         Left            =   1470
         TabIndex        =   52
         Top             =   4560
         Width           =   9060
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Losing Lives:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2760
         TabIndex        =   51
         Top             =   4200
         Width           =   6480
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":27707
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1080
         Left            =   1470
         TabIndex        =   50
         Top             =   2880
         Width           =   9060
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Object of the Game:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2760
         TabIndex        =   49
         Top             =   2520
         Width           =   6480
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NumPad 3: Down-Right"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2550
         TabIndex        =   48
         Top             =   7320
         Width           =   3240
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NumPad 1: Down-Left"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2655
         TabIndex        =   47
         Top             =   6960
         Width           =   3060
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NumPad 9: Up-Right"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2775
         TabIndex        =   46
         Top             =   6600
         Width           =   2820
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   555
         Left            =   5475
         TabIndex        =   45
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keyboard Controls (Turn NumLock On):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2760
         TabIndex        =   44
         Top             =   5880
         Width           =   6480
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NumPad 7: Up-Left"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2880
         TabIndex        =   43
         Top             =   6240
         Width           =   2640
      End
   End
   Begin VB.PictureBox picAbout 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF8080&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin Project1.Button Button2 
         Height          =   495
         Left            =   9960
         TabIndex        =   37
         Top             =   8280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OK"
         Shape           =   0
         FillColor       =   12582912
         FillColorMouseOver=   12582912
         FillColorMouseDown=   16711680
         ForeColorInvert =   0   'False
         FontChangeMouseDown=   0
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visit IsoEngine's web site at www.firstproductions.com/isoengine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2760
         TabIndex        =   40
         Top             =   3480
         Width           =   6480
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3D Snake is a sample that comes with IsoEngine SDK."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2760
         TabIndex        =   39
         Top             =   2520
         Width           =   6480
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   555
         Left            =   5310
         TabIndex        =   38
         Top             =   1920
         Width           =   1380
      End
   End
   Begin VB.PictureBox picGame 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   7680
      Left            =   240
      ScaleHeight     =   512
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   768
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   11520
   End
   Begin Project1.Button btnStop 
      Height          =   495
      Left            =   9960
      TabIndex        =   7
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Stop"
      Shape           =   0
      FillColor       =   12582912
      FillColorMouseOver=   12582912
      FillColorMouseDown=   16711680
      ForeColorInvert =   0   'False
      FontChangeMouseDown=   0
   End
   Begin VB.PictureBox Focus 
      Height          =   255
      Left            =   600
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   -15000
      Width           =   135
   End
   Begin Project1.Button btnMenu 
      Height          =   735
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   2213
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Play"
      Shape           =   0
      FillColor       =   16744576
      FillColorMouseOver=   16711680
      FillColorMouseDown=   12582912
      ForeColorInvert =   0   'False
      FontChangeMouseDown=   0
   End
   Begin Project1.Button btnMenu 
      Height          =   735
      Index           =   1
      Left            =   5160
      TabIndex        =   1
      Top             =   3173
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Options"
      Shape           =   0
      FillColor       =   16744576
      FillColorMouseOver=   16711680
      FillColorMouseDown=   12582912
      ForeColorInvert =   0   'False
      FontChangeMouseDown=   0
   End
   Begin Project1.Button btnMenu 
      Height          =   735
      Index           =   2
      Left            =   5160
      TabIndex        =   2
      Top             =   4133
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Help"
      Shape           =   0
      FillColor       =   16744576
      FillColorMouseOver=   16711680
      FillColorMouseDown=   12582912
      ForeColorInvert =   0   'False
      FontChangeMouseDown=   0
   End
   Begin Project1.Button btnMenu 
      Height          =   735
      Index           =   3
      Left            =   5160
      TabIndex        =   3
      Top             =   5093
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "About"
      Shape           =   0
      FillColor       =   16744576
      FillColorMouseOver=   16711680
      FillColorMouseDown=   12582912
      ForeColorInvert =   0   'False
      FontChangeMouseDown=   0
   End
   Begin Project1.Button btnMenu 
      Height          =   735
      Index           =   4
      Left            =   5160
      TabIndex        =   4
      Top             =   6053
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      Shape           =   0
      FillColor       =   16744576
      FillColorMouseOver=   16711680
      FillColorMouseDown=   12582912
      ForeColorInvert =   0   'False
      FontChangeMouseDown=   0
   End
   Begin Project1.Button btnPause 
      Height          =   495
      Left            =   8160
      TabIndex        =   8
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pause"
      Shape           =   0
      FillColor       =   12582912
      FillColorMouseOver=   12582912
      FillColorMouseDown=   16711680
      ForeColorInvert =   0   'False
      FontChangeMouseDown=   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnMenu_Click(index As Integer)
    Select Case index
        Case 0
            picGame.Visible = True
            btnPause.Visible = True
            btnStop.Visible = True
            ResetGame
            RenderLoop
        Case 1
            picOptions.Picture = Picture
            picOptions.Visible = True
        Case 2
            picHelp.Picture = Picture
            picHelp.Visible = True
        Case 3
            picAbout.Picture = Picture
            picAbout.Visible = True
        Case 4
            UnloadGame
            Helper.TASKBAR_Show
            End
    End Select
End Sub

Private Sub btnPause_Click()
    Pause = Not Pause
    If btnPause.Caption = "Pause" Then
        btnPause.Caption = "Resume"
    Else
        btnPause.Caption = "Pause"
    End If
End Sub

Private Sub btnPause_KeyDown(KeyCode As Integer, Shift As Integer)
    picGame_KeyDown KeyCode, Shift
End Sub

Public Sub btnStop_Click()
    picGame.Visible = False
    btnStop.Visible = False
    btnPause.Visible = False
    StopGame = True
End Sub

Private Sub btnStop_KeyDown(KeyCode As Integer, Shift As Integer)
    picGame_KeyDown KeyCode, Shift
End Sub

Private Sub Button1_Click()
    StartAtLevel = HScroll5.Value
    FruitsPerLevel = HScroll4.Value
    SoundVolume = HScroll3.Value * 300 - 3600
    MusicVolume = HScroll2.Value * 300 - 3600
    SnakeSpeed = (21 - HScroll1.Value) * 0.01
    picOptions.Visible = False
End Sub

Private Sub Button2_Click()
    picAbout.Visible = False
End Sub

Private Sub Button3_Click()
    picHelp.Visible = False
End Sub

Private Sub Form_Load()
    SnakeSpeed = 0.06
    MusicVolume = -1300
    SoundVolume = 0
    FruitsPerLevel = 10
    StartAtLevel = 0
    Set picLevel = LoadPicture(App.Path & "\Graphics\lvl0.jpg")
    LoadIsoEngine
    Helper.TASKBAR_Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Focus.SetFocus
    On Error Resume Next
    frmMain.picGame.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadGame
    Helper.TASKBAR_Show
    End
End Sub

Private Sub HScroll1_Change()
    Label14 = HScroll1.Value
    
    'Sets the tempo based on the speed of the snake
    If Label14 = 20 Then
        Music.Tempo = 125
    ElseIf Label14 > 15 Then
        Music.Tempo = 112
    ElseIf Label14 > 5 Then
        Music.Tempo = 100
    ElseIf Label14 > 1 Then
        Music.Tempo = 88
    Else
        Music.Tempo = 75
    End If
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub HScroll2_Change()
    Label15 = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
    HScroll2_Change
End Sub

Private Sub HScroll3_Change()
    Label16 = HScroll3.Value
End Sub

Private Sub HScroll3_Scroll()
    HScroll3_Change
End Sub

Private Sub HScroll4_Change()
    Label17 = HScroll4.Value
End Sub

Private Sub HScroll4_Scroll()
    HScroll4_Change
End Sub

Private Sub HScroll5_Change()
    Label18 = HScroll5.Value
    Set picLevel = LoadPicture(App.Path & "\Graphics\lvl" & Label18 & ".jpg")
End Sub

Private Sub HScroll5_Scroll()
    HScroll5_Change
End Sub

Private Sub picGame_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyNumpad7 'Up-Left
            If Direction(UBound(Direction)).X = 2 And Direction(UBound(Direction)).Y = 1 Then Exit Sub
            ReDim Preserve Direction(UBound(Direction) + 1)
            Direction(UBound(Direction)).X = -2
            Direction(UBound(Direction)).Y = -1
        Case vbKeyNumpad9 'Up-Right
            If Direction(UBound(Direction)).X = -2 And Direction(UBound(Direction)).Y = 1 Then Exit Sub
            ReDim Preserve Direction(UBound(Direction) + 1)
            Direction(UBound(Direction)).X = 2
            Direction(UBound(Direction)).Y = -1
        Case vbKeyNumpad3 'Down-Right
            If Direction(UBound(Direction)).X = -2 And Direction(UBound(Direction)).Y = -1 Then Exit Sub
            ReDim Preserve Direction(UBound(Direction) + 1)
            Direction(UBound(Direction)).X = 2
            Direction(UBound(Direction)).Y = 1
        Case vbKeyNumpad1 'Down-Left
            If Direction(UBound(Direction)).X = 2 And Direction(UBound(Direction)).Y = -1 Then Exit Sub
            ReDim Preserve Direction(UBound(Direction) + 1)
            Direction(UBound(Direction)).X = -2
            Direction(UBound(Direction)).Y = 1
        Case vbKeyRight
            ReDim Preserve Direction(UBound(Direction) + 1)
            If Direction(UBound(Direction) - 1).X = -2 And Direction(UBound(Direction) - 1).Y = -1 Then
                Direction(UBound(Direction)).X = 2
                Direction(UBound(Direction)).Y = -1
            ElseIf Direction(UBound(Direction) - 1).X = 2 And Direction(UBound(Direction) - 1).Y = -1 Then
                Direction(UBound(Direction)).X = 2
                Direction(UBound(Direction)).Y = 1
            ElseIf Direction(UBound(Direction) - 1).X = 2 And Direction(UBound(Direction) - 1).Y = 1 Then
                Direction(UBound(Direction)).X = -2
                Direction(UBound(Direction)).Y = 1
            Else
                Direction(UBound(Direction)).X = -2
                Direction(UBound(Direction)).Y = -1
            End If
        Case vbKeyLeft
            ReDim Preserve Direction(UBound(Direction) + 1)
            If Direction(UBound(Direction) - 1).X = -2 And Direction(UBound(Direction) - 1).Y = -1 Then
                Direction(UBound(Direction)).X = -2
                Direction(UBound(Direction)).Y = 1
            ElseIf Direction(UBound(Direction) - 1).X = 2 And Direction(UBound(Direction) - 1).Y = -1 Then
                Direction(UBound(Direction)).X = -2
                Direction(UBound(Direction)).Y = -1
            ElseIf Direction(UBound(Direction) - 1).X = 2 And Direction(UBound(Direction) - 1).Y = 1 Then
                Direction(UBound(Direction)).X = 2
                Direction(UBound(Direction)).Y = -1
            Else
                Direction(UBound(Direction)).X = 2
                Direction(UBound(Direction)).Y = 1
            End If
        Case vbKeyP
            btnPause_Click
        Case vbKeyEscape
            btnStop_Click
    End Select
End Sub

Private Sub Picture2_Click()

End Sub
