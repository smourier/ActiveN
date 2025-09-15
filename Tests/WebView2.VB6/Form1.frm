VERSION 5.00
Object = "{D66AD502-5443-4B3F-93E4-EC0975FD3408}#1.0#0"; "ActiveN.Samples.WebView2.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin ActiveNSamplesWebView2LibCtl.WebView2Control WebView2Control1 
      Height          =   6255
      Left            =   360
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   2
      Top             =   960
      Width           =   10695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   495
      Left            =   10080
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "https://www.simonmourier.com"
      Top             =   240
      Width           =   9615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    WebView2Control1.Navigate Text1.Text
End Sub
