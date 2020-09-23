Attribute VB_Name = "IntersectHM"
' this API is reported for comparison, but not used at all
'Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
'Public Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public MaxRETCANGLES As Integer  '(for demo, suggested value = 5, for fun up to 250), change this to limit max number of allowed rectangle that intersect eachother

Public Function IntersectHomeMade(ByVal DestRectobj As Object, ByRef SourceRectobj() As Object, ByVal SourceRectCount As Integer) As Long
   'inputs: more than 2 source rectangular objects (SourceRectobj is a matrix)
   'outputs: - return of function IntersectHomeMade = 0 --> source objects DO NOT intersect
   '         - return of function IntersectHomeMade <> 0 --> source objects DO intersect
   '         - DestRectobj --> intersection rectangle obj
   'NOTE: to be seen, DestRectobj must be contaibed in a different container as the two source objects
   'for example: the 2 sources on the form and the destination object into a picture
   'EXAMPLE CALL: [drop down 2 labels in form1 and name them Source1Label and Source2Label
   '               then a picture
   '              ret=IntersectHomeMade(Form1.IntLabel,Form1.Source1Label,Form1.Source2Label)
   '              if ret=0 then '2 source labels intersect
   '              else          '2 source labels DO NOT intersect
   '              endif
   
   
   Dim verline(1 To 4) As Single
   Dim horline(1 To 4) As Single
   Dim IntersectPoint(1 To 16) As POINTAPI
   Dim IntersectPointGood(1 To 4) As POINTAPI
   
   'define a fictious object (not existent,
   'is not an object either, just a set of
   'variables) that we'll call
   'fictiousintersectobj, whose 'properties'
   'are:
   Dim FictiousIntersectObjLEFT As Integer
   Dim FictiousIntersectObjTOP As Integer
   Dim FictiousIntersectObjWIDTH As Integer
   Dim FictiousIntersectObjHEIGHT As Integer

   'let's initialize it with
   'SourceRectobj(1) properties
   FictiousIntersectObjLEFT = SourceRectobj(1).Left
   FictiousIntersectObjTOP = SourceRectobj(1).Top
   FictiousIntersectObjWIDTH = SourceRectobj(1).Width
   FictiousIntersectObjHEIGHT = SourceRectobj(1).Height

   
   'this is inspired to Tarek Said algorithm
   'generalized to more than 2 rectangles
   For aa = 2 To SourceRectCount
      TheyIntersect = 0
      Select Case SourceRectobj(aa).Left
         Case FictiousIntersectObjLEFT - SourceRectobj(aa).Width To FictiousIntersectObjLEFT + FictiousIntersectObjWIDTH
             Select Case SourceRectobj(aa).Top
                Case FictiousIntersectObjTOP - SourceRectobj(aa).Height To FictiousIntersectObjTOP + FictiousIntersectObjHEIGHT
                   'make SourceRectobj(1)= intersec rect
                   GoSub findintrect
                   TheyIntersect = 1
                Case Else
                   TheyIntersect = 0
                   Exit For
             End Select
         Case Else
            TheyIntersect = 0
            Exit For
      End Select
   Next

   If TheyIntersect = 0 Then 'don't intersect, exit
      DestRectobj.Width = 0
      DestRectobj.Height = 0
      DestRectobj.Visible = False
      DestRectobj.Refresh
   Else                          'draw intersect rect
      'DestRectobj may now be returned (intersection rectangle)
      DestRectobj.Left = minx
      DestRectobj.Top = miny
      DestRectobj.Width = maxx - minx
      DestRectobj.Height = maxy - miny
      DestRectobj.Visible = True
      DestRectobj.Refresh
   End If
   IntersectHomeMade = TheyIntersect


Exit Function


findintrect:
   
   'this is my algorithm for obtainiung the intersection rectangle
   'find the intersect rectangle (its upperleft and downright points)
   'the rectangles intersect
   'find the 16 points (real and virtual intersections)
   'basing on these possible coordinates:
   verline(1) = FictiousIntersectObjLEFT
   verline(2) = FictiousIntersectObjLEFT + FictiousIntersectObjWIDTH
   verline(3) = SourceRectobj(aa).Left
   verline(4) = SourceRectobj(aa).Left + SourceRectobj(aa).Width
   horline(1) = FictiousIntersectObjTOP
   horline(2) = FictiousIntersectObjTOP + FictiousIntersectObjHEIGHT
   horline(3) = SourceRectobj(aa).Top
   horline(4) = SourceRectobj(aa).Top + SourceRectobj(aa).Height
   'here are the 16 points
   For a = 1 To 4
      For b = 1 To 4
         IntersectPoint((a - 1) * 4 + b).x = verline(a)
         IntersectPoint((a - 1) * 4 + b).y = horline(b)
      Next
   Next
   
   'now, only 4 of the above 16 points are the ones of
   'the intersection rectangle. they must belong to both
   'rectangles.
   'find them:

   inc = 0
   For a = 1 To 16
      Select Case IntersectPoint(a).x
         Case verline(1) To verline(2) 'part of SourceRectobj1
            Select Case IntersectPoint(a).x
               Case verline(3) To verline(4) ' and part of SourceRectobj2
                  Select Case IntersectPoint(a).y
                     Case horline(1) To horline(2) 'part of SourceRectobj1
                        Select Case IntersectPoint(a).y
                           Case horline(3) To horline(4) 'and part of SourceRectobj2
                              'this is one of four points of intersect rectangle
                              'collect it
                              inc = inc + 1
                              'prevent possible mistaken coordinates
                              'caused by the high dynamics
                              '(move rect 1 and compute intersect
                              'at the same time)
                              If inc > 4 Then Exit For
                              IntersectPointGood(inc).x = IntersectPoint(a).x
                              IntersectPointGood(inc).y = IntersectPoint(a).y
                        End Select
                  End Select
            End Select
      End Select
   Next


   'find line bounds of intersect rectangle
   minx = Screen.Width - 1
   miny = Screen.Height - 1
   maxx = 0
   maxy = 0
   
   For a = 1 To 4
      If IntersectPointGood(a).x > maxx Then maxx = IntersectPointGood(a).x
      If IntersectPointGood(a).x < minx Then minx = IntersectPointGood(a).x
      If IntersectPointGood(a).y > maxy Then maxy = IntersectPointGood(a).y
      If IntersectPointGood(a).y < miny Then miny = IntersectPointGood(a).y
   Next
   
   
   'fictious intersect object size and position is now
   FictiousIntersectObjLEFT = minx
   FictiousIntersectObjTOP = miny
   FictiousIntersectObjWIDTH = maxx - minx
   FictiousIntersectObjHEIGHT = maxy - miny
   
Return

End Function


Public Sub Pause(ByVal milliseconds As Double) 'milliseconds
    Start = Timer * 1000 ' Set start time.
    Do While Timer * 1000 < Start + milliseconds
        DoEvents    ' Yield to other processes.
        If Timer * 1000 < Start Then 'trepassing midnight
          Start = Start - 24! * 60 * 60 * 1000
        End If
    Loop
End Sub
