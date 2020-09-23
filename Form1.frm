VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Y! Movies"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   4050
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   3375
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   570
      ScaleWidth      =   2730
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   2760
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.TextBox Text2 
      Height          =   5325
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "Form1.frx":014A
      Top             =   4365
      Width           =   7755
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2070
      TabIndex        =   6
      Text            =   "http://www.planetsourcecode.com/vb"
      Top             =   90
      Width           =   5145
   End
   Begin VB.ListBox List4 
      Height          =   1425
      Left            =   4770
      TabIndex        =   5
      Top             =   2790
      Width           =   3075
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   1935
      TabIndex        =   4
      Top             =   2790
      Width           =   2805
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   135
      TabIndex        =   3
      Top             =   2790
      Width           =   1680
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   90
      TabIndex        =   1
      Top             =   675
      Width           =   7755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Go =>"
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   1680
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   450
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 

Private Type Pointapi
   X As Long
   Y As Long
End Type


Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, ByRef lpPoint As Pointapi) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Pointapi) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long) As Long




Dim old_hwnd As Long
Dim ttip_cnt As Long
  
Dim WithEvents HTML As cHtmlDoc
Attribute HTML.VB_VarHelpID = -1
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long








  

Private Sub Command1_Click()
   
  List1.Clear: List2.Clear: List3.Clear: List4.Clear
  HTML.document_from_url (Text1), , True, True, True, True
 
End Sub

Private Sub Form_Load()
 
   Set HTML = New cHtmlDoc
   SetParent Picture1.hwnd, GetDesktopWindow
   Timer1.Interval = 2000
   Timer1 = True
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Timer1 = False
    Set HTML = Nothing
    'since we made the desktop the parent of picture1
    'in form load..we need to explicitly unload it here
    'by setting its parent back to this form
    SetParent Picture1.hwnd, hwnd
    
End Sub

 

Private Sub HTML_docCreateProgress(progress As String)

  lblProgress = progress
  
End Sub

Private Sub HTML_docCreationFailed()
 
  MsgBox "Document Creation Failed !"
  
End Sub

Private Sub HTML_docReady(htmlDoc As MSHTML.HTMLDocument)
  Dim i As Integer
  
  lblProgress = "DONE!"
  
  For i = 0 To htmlDoc.All.Length - 1
   ' Debug.Print htmlDoc.All(i).innerHTML
  Next i
 
End Sub

Private Sub HTML_bodyReady(htmlBody As MSHTML.htmlBody)
 
 Text2 = htmlBody.innerText
  
End Sub

Private Sub HTML_imagesReady(aImage As MSHTML.HTMLImg)

  List4.AddItem aImage.fileSize & vbTab & aImage.href

End Sub

Private Sub HTML_inputElementsReady(aInputElement As MSHTML.HTMLInputElement)

  List2.AddItem aInputElement.Value

End Sub

Private Sub HTML_linksReady(aLink As MSHTML.HTMLAnchorElement)
  
  List1.AddItem aLink.href
  
End Sub

Private Sub HTML_tableCellsReady(aTableCell As MSHTML.HTMLTableCell)

 List3.AddItem aTableCell.cellIndex & vbTab & aTableCell.innerText

End Sub
 
 
'because autosize=True, this event also
'is a resize trigger for the label
Private Sub Label1_Change()
 
 Picture1.Width = Label1.Width + 250
 Picture1.Height = Label1.Height + 50

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Picture1.Visible = False

End Sub

Private Sub Text1_DblClick()
 
  With Text1
      .Text = "http://www..com"
      .SelStart = 11
  End With
 
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then KeyAscii = 0: Command1_Click
  
End Sub

Private Sub Timer1_Timer()

  Dim lHwnd        As Long
  Dim new_hwnd     As Long
  Dim pix_wid As Long, pix_hei   As Long
  Dim pt           As Pointapi
  Static bShown    As Boolean
  
  GetCursorPos pt
  new_hwnd = WindowFromPoint(pt.X, pt.Y)
 
 
  If old_hwnd <> new_hwnd Then 'means the mouse moved to different control
      old_hwnd = new_hwnd 'store the new window
      
      With Picture1 'position the piturebox using api friendly coods
         pix_wid = .Width / Screen.TwipsPerPixelX
         pix_hei = .Height / Screen.TwipsPerPixelY
         MoveWindow .hwnd, pt.X + 10, pt.Y, pix_wid, pix_hei, True
         SetWindowPos Picture1.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
      End With
   
       Select Case new_hwnd 'change the caption for different listboxes
          Case Is = List1.hwnd
              Label1.Caption = "All of the documents links" & vbCrLf & _
                               "will be displayed here"
              Picture1.Visible = True
              
          Case Is = List2.hwnd
               Label1.Caption = "All of the documents input" & vbCrLf & _
                                "elements values will be displayed here"
               Picture1.Visible = True
               
          Case Is = List3.hwnd
               Label1.Caption = "All of the documents table cells .cellIndex" & vbCrLf & _
                                "& .innerText will be displayed here"
               Picture1.Visible = True
              
          Case Is = List4.hwnd
               Label1.Caption = "All of the documents images size (in bytes)" & vbCrLf & _
                                "and href will be displayed here"
               Picture1.Visible = True
               
           Case Is = Text2.hwnd
               Label1.Caption = "The "".innerHTML"" property of the" & vbCrLf & _
                                "documents body object is displayed here" & vbCrLf & _
                                "Makes creating a text only browser pretty" & vbCrLf & _
                                "damned easy  ___oo0o____( . ) ( . )____o0oo____"
               Picture1.Visible = True
               
           Case Is = Text1.hwnd
               Label1.Caption = "Double click to reset the web address"
               Picture1.Visible = True
               
       End Select

       
    Else
           Picture1.Visible = False
    End If
    
  
End Sub
