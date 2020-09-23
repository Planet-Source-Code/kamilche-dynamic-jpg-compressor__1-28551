VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Convert Bitmap to JPG"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCompression 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Left            =   2595
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3825
      Width           =   3750
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   420
      Left            =   15
      TabIndex        =   1
      Top             =   4065
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   741
      _Version        =   393216
      Max             =   100
      SelStart        =   1
      TickFrequency   =   10
      Value           =   1
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   2670
      Left            =   0
      ScaleHeight     =   2610
      ScaleWidth      =   6285
      TabIndex        =   0
      Top             =   0
      Width           =   6345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BMP As String
Private JPG As String
Private Quality As Long

Private Declare Function BMPToJPG Lib "converter.dll" (ByVal InputFilename As String, ByVal OutputFilename As String, ByVal Quality As Long) As Integer

Private Sub Form_Load()
    BMP = App.Path & "\picture.bmp"
    JPG = App.Path & "\picture.jpg"
    Slider1.Value = 0
End Sub


Private Sub Slider1_Change()
    Quality = 100 - Slider1.Value
    If Quality < 1 Then
        Quality = 1
    ElseIf Quality > 100 Then
        Quality = 100
    End If
    If BMPToJPG(BMP, JPG, Quality) = 0 Then
        Picture1.Picture = LoadPicture(JPG)
        txtCompression.Text = "Compression: " & Slider1.Value & vbCrLf & _
        "BMP size: " & FileLen(BMP) & vbCrLf & _
        "JPG size: " & FileLen(JPG)
    Else
        MsgBox "Error in converting! Make sure the input filename exists, and is a valid BMP file.", vbCritical
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        Picture1.Height = ScaleHeight - txtCompression.Height
        txtCompression.Top = ScaleHeight - txtCompression.Height
        Slider1.Top = txtCompression.Top + 100
    End If
End Sub

Private Sub Slider1_Scroll()
    Slider1_Change
End Sub
