VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.v12.0.1.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4980
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmd 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock sckAdmin 
      Left            =   5880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   0
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "·þÎñÆ÷¶Ë¿Ú£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "·þÎñÆ÷µØÖ·£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Êý¾Ý¿âÃÜÂë£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Êý¾Ý¿âÕËºÅ£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim Descriptions As SkinDescriptions
    Set Descriptions = SkinFramework.EnumerateSkinDirectory(App.Path + "\", True)
    Dim Des As SkinDescription
    For Each Des In Descriptions
        Debug.Print Des.Name & " - " & Des.Path
    Next
    SkinFramework.LoadSkin App.Path & "\Styles\iTunes", "Normalitunes.ini"
    SkinFramework.ApplyWindow Me.hWnd
    SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    
End Sub
