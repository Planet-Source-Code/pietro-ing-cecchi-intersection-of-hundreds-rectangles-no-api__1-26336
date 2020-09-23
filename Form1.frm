VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Home Made Intersect (of hundreds  rectangles)"
   ClientHeight    =   4020
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton AddRemove 
      Caption         =   "remove a rectangle"
      Height          =   375
      Index           =   1
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton AddRemove 
      Caption         =   "add a rectangle"
      Height          =   375
      Index           =   0
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton ChangeShapes 
      Caption         =   "change size and color"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   960
      Width           =   2655
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "moveable object, click down and drag to move"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "2"
         Height          =   735
         Index           =   1
         Left            =   840
         TabIndex        =   6
         ToolTipText     =   "moveable object, click down and drag to move"
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   2880
      ScaleHeight     =   2235
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   960
      Width           =   2655
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "do they intersect?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   4
      Top             =   600
      Width           =   5445
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "directions"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Home made Intersect is more powerful than the Api call IntersectRect (only 2 rectangles) "
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "ABOVE:    the intersection rectangle"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      ToolTipText     =   "No more need of Api call, with this powerful home made intersect routine!"
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Menu menufile 
      Caption         =   ".  File  ."
      Begin VB.Menu menuhighbound 
         Caption         =   "high bound change"
      End
      Begin VB.Menu menuline0 
         Caption         =   "-"
      End
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
      End
      Begin VB.Menu menuline1 
         Caption         =   "-"
      End
      Begin VB.Menu menucancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu menuabout 
      Caption         =   ".  About  ."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xsave As Single
Dim Ysave As Single
Dim MouseIsDown As Boolean

Dim NumberOfSourceRectangles As Integer

Const colminR = 100
Const colrndR = 255 - colminR
Const colminG = 50
Const colrndG = 255 - colminG
Const colminB = 100
Const colrndB = 255 - colminB


Private Sub AddRemove_Click(Index As Integer)
   Select Case Index
      Case 0 'add a rectangle
         If NumberOfSourceRectangles + 1 <= MaxRETCANGLES Then
            Load Label7(NumberOfSourceRectangles)
            With Label7(NumberOfSourceRectangles)
               .Caption = NumberOfSourceRectangles + 1
               .Width = Rnd * (1544 - 255) + 255
               .Height = Rnd * (1544 - 255) + 255
               .Move (Picture2.Width - .Width) / 2, (Picture2.Height - .Height) / 2
               RandomColor Label7(NumberOfSourceRectangles)
               .ZOrder
               .Visible = True
            End With
            NumberOfSourceRectangles = NumberOfSourceRectangles + 1
            ManageIntersect
         Else
            captsave = AddRemove(0).Caption
            colsave = AddRemove(0).BackColor
            AddRemove(0).BackColor = vbYellow
            AddRemove(0).Caption = "hight bound:" & MaxRETCANGLES
            Pause 1000
            AddRemove(0).Caption = captsave
            AddRemove(0).BackColor = colsave
         End If
      Case 1 'remove a rectangle
         If NumberOfSourceRectangles - 1 >= 2 Then
            Unload Label7(NumberOfSourceRectangles - 1)
            NumberOfSourceRectangles = NumberOfSourceRectangles - 1
            ManageIntersect
         Else
            captsave = AddRemove(1).Caption
            colsave = AddRemove(1).BackColor
            AddRemove(1).BackColor = vbYellow
            AddRemove(1).Caption = "low bound: 2"
            Pause 1000
            AddRemove(1).Caption = captsave
            AddRemove(1).BackColor = colsave
         End If
   End Select
End Sub

Private Sub ChangeShapes_Click()
   Randomize Timer
   For Each Control In Controls
       On Error Resume Next
       If Control.Container.Name = "Picture2" Then
          If Err > 0 Then Err.Clear: On Error GoTo 0: GoTo skipthis
          On Error GoTo 0
          Control.Width = Rnd * (1544 - 255) + 255
          Control.Height = Rnd * (1544 - 255) + 255
          RandomColor Control
       End If
skipthis:
   Next
   ManageIntersect
End Sub

Private Sub Form_Load()
   MaxRETCANGLES = 5  '(for demo, suggested value = 5, for fun up to 250), change this to limit max number of allowed rectangle that intersect eachother
   NumberOfSourceRectangles = 2 'passed to IntersectHomeMade
   RandomColor Label7(0)
   RandomColor Label7(1)
   Label4.Caption = "DIRECTIONS: click down and drag any rectangle to create a common intersection. DblClk to raise/lower"
   menuabout_Click
End Sub
Private Sub RandomColor(ByVal obj As Object)
   With obj
      Randomize Timer
      
      'tricky way of obtaining bright random colors
      red = colminR + Rnd * colrndR
      green = colminG + Rnd * colrndG
      blue = colminB + Rnd * colrndB
      Max = 0
      If green > red Then Max = green Else Max = red
      If blue > Max Then Max = blue
      clamp = 255
      If red = Max Then red = clamp
      If green = Max Then green = clamp
      If blue = Max Then blue = clamp

      If red + green + blue < 2 * (colminR + colminG + colminB) Then
         .ForeColor = vbWhite
      Else
         .ForeColor = vbBlack
      End If
      
      .BackColor = RGB(red, green, blue)
   
   End With
End Sub

Private Sub Label7_DblClick(Index As Integer)
   Static indsave As Integer
   Static ff As Boolean
   If indsave = Index Then
      ff = Not ff
   Else
      ff = True
   End If
   If ff Then
      Label7(Index).ZOrder
   Else
      Label7(Index).ZOrder 1
   End If
   indsave = Index
End Sub

Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   MouseIsDown = True
   Xsave = x
   Ysave = y
End Sub

Private Sub Label7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not MouseIsDown Then Exit Sub
   With Label7(Index)
      .Move .Left + (x - Xsave), .Top + (y - Ysave)
   End With
   ManageIntersect
End Sub

Private Sub Label7_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   MouseIsDown = False
End Sub

Private Sub ManageIntersect()
   ReDim sourceobj(1 To NumberOfSourceRectangles) As Object
   For a = 1 To NumberOfSourceRectangles
      Set sourceobj(a) = Label7(a - 1)
   Next
   sourceobjcount = NumberOfSourceRectangles
   s1 = "1 to " & sourceobjcount & " "
   ssyes = s1 & " DO  intersect"
   ssno = s1 & " DO NOT  intersect"
   Label3.Caption = IIf((IntersectHomeMade(Label5, sourceobj(), sourceobjcount)) > 0, ssyes, ssno)
End Sub

Private Sub menuhighbound_Click()
again:
   response = InputBox("You can change here the upper bound for rectangles addition." & vbNewLine & _
                       "Possible range is 2 to 250 rectangles.", "Change max number of rectangles", MaxRETCANGLES)
   If response = "" Then Exit Sub
   response = CInt(response)
   Select Case response
      Case 2 To 250
         MaxRETCANGLES = response
      Case Else
         DoEvents
         GoTo again
   End Select
End Sub

Private Sub menuexit_Click()
   Unload Me
End Sub

Private Sub menuabout_Click()
   MsgBox "Curious of how to get the intersection of hundreds rectangles almost instantly and without any API call?" & vbNewLine & vbNewLine & _
          "Then come and see!..." & vbNewLine & vbNewLine & _
          "          Arabian and italian algorithms in play!", vbOKOnly + vbInformation, "Homemade Intersect"
          
End Sub
