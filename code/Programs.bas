Attribute VB_Name = "Programs"
''''''''' (C) 2020 VZ Home Experiments Vladimir Zhbanko https://vladdsm.github.io/myblog_attempt/
''''''''' VBA code to make work with Excel User Forms easier
''''''''' More time to spend on more interesting stuff.
''''''''' For donations and support: https://www.paypal.me/Zhbanko
Option Explicit
Dim wb As Workbook
Dim wshRep As Worksheet: Dim wshPln As Worksheet: Dim wshPiv As Worksheet

'========================================
'PASTE PICTURE TO CELL
'========================================
' 2020-04-01
' @Description This program [Sub] gets picture path and the row number, it will place picture into the excel cell
' @Detail Program will perform picture transformation to compress the picture
'
' @param picPath - string, path to the picture
' @param iRow - integer, row in the worksheet

Sub PastePicture(picPath, iRow)

  ' resize row height first
  Worksheets("Report").Rows(iRow).RowHeight = 79

      With Worksheets("Report").Pictures.Insert(picPath)
      
        With .ShapeRange
            .LockAspectRatio = msoTrue
            .Width = 90         'width of the picture
            .Height = 75        'height of the picture
        End With
        ' define where to place the picture in the cell
        .Left = Worksheets("Report").CELLS(iRow, 15).Left + 2
        .Top = Worksheets("Report").CELLS(iRow, 15).Top + 2
        .Placement = 1
        .PrintObject = True
        .Name = "Sample" & iRow    ' use .Name property to name the picture with known name
      
        ' optimize RAM usage by keeping the picture in the cell, not linked to folder source
        ' using the "known" name we perform operation on the picture
        With ActiveSheet.Shapes.Range(Array("Sample" & iRow)).Select
          Selection.Cut
          CELLS(iRow, 15).Select
          ActiveSheet.Pictures.Paste.Select
          ' method to move the Shape
          Selection.ShapeRange.IncrementLeft 2
          Selection.ShapeRange.IncrementTop 2
          CELLS(iRow, 15).Select
        End With
       
    End With
   
End Sub

'========================================
'UPDATE USER FORM INPUTS
'========================================
' 2020-04-01
' @Description This program [Sub] brings the input information to the User Form
' information is found using 'iRow' argument that represents worksheet row
' Program is also checking if detailed information is already present on the worksheet Report
' if that is so, then such information will be pulled from the worksheet Report
' such functionality is created using unique 'id' that must match on both Sheets
'
' @param iRow - integer, row in the worksheet
'
Sub UpdateInputs(iRow)
' --------------------
' Define variables needed (typically all the fields name on both sheets)
' --------------------
Dim SheetName As String: Dim itemIDval As String: Dim isMatchingIdOnPlanning As Boolean: Dim isMatchingIdOnReport As Boolean
Dim j As Long: Dim lRow As Integer: Dim LocationMatchingID As Long
' Variables used
Dim Item As String: Dim Owner As String: Dim STGoal As String: Dim LTGoal As String
Dim Activity As String: Dim Status As String: Dim Situation As String: Dim Comments As String
Dim StatusKey As Integer: Dim SituationKey As Integer
Dim StDate As String: Dim EnDate As String
Dim Expenses As Double: Dim HoursSpend As Long: Dim ValueAdded As Double: Dim picPath As String
Set wb = ThisWorkbook: Set wshRep = wb.Worksheets("Report"): Set wshPln = wb.Worksheets("Planning")
' --------------------
' Initialize variables
' --------------------
' Determine the worksheet where user is invoking the UserForm
SheetName = ActiveSheet.Name
' Determine Unique id on the worksheet (must be present before user clicks!)
Item = Range("A" & iRow).Value

' Find if matching data [id] are present on the different worksheet Report/Planning
isMatchingIdOnPlanning = isExistingID(Item, "Planning")
isMatchingIdOnReport = isExistingID(Item, "Report")

' *
' [User clicked from worksheet planning] and such record is present on Report
If SheetName = "Planning" And isMatchingIdOnReport = True Then
    ' get data from the WS Planning
    Owner = wshPln.Range("B" & iRow).Value
    STGoal = wshPln.Range("C" & iRow).Value
    LTGoal = wshPln.Range("D" & iRow).Value
    Activity = wshPln.Range("E" & iRow).Value
    Status = wshPln.Range("F" & iRow).Value
    Situation = wshPln.Range("G" & iRow).Value
    Comments = wshPln.Range("H" & iRow).Value
    
    ' find row location of Item on WS Report
    lRow = getRowID(Item, "Report")
    
    ' get remaining data from the WS Report
    StDate = wshRep.Range("B" & iRow).Value
    If StDate = "" Then
        StDate = CStr(Format(Date + 1, "dd/mm/yyyy")) 'help user with default data entry!
    End If
    EnDate = wshRep.Range("C" & iRow).Value
    If EnDate = "" Then
        EnDate = CStr(Format(Date + 10, "dd/mm/yyyy")) 'help user with default data entry!
    End If
    Expenses = wshRep.Range("K" & lRow).Value
    HoursSpend = wshRep.Range("L" & lRow).Value
    ValueAdded = wshRep.Range("M" & lRow).Value
    picPath = wshRep.Range("N" & lRow).Value
    
' [User clicked from worksheet planning] and such record is not present on Report
ElseIf SheetName = "Planning" And isMatchingIdOnReport = False Then
    
    ' get data from the WS Planning
    Owner = wshPln.Range("B" & iRow).Value
    STGoal = wshPln.Range("C" & iRow).Value
    LTGoal = wshPln.Range("D" & iRow).Value
    Activity = wshPln.Range("E" & iRow).Value
    Status = wshPln.Range("F" & iRow).Value
    Situation = wshPln.Range("G" & iRow).Value
    Comments = wshPln.Range("H" & iRow).Value
    
    ' Step 2. Guess some default values?
    StDate = CStr(Format(Date + 1, "dd/mm/yyyy"))
    EnDate = CStr(Format(Date + 10, "dd/mm/yyyy"))

' [User clicked from worksheet Report] and such record is present on WS Planning
ElseIf SheetName = "Report" And isMatchingIdOnPlanning = True Then

    StDate = wshRep.Range("B" & iRow).Value
    If StDate = "" Then
        StDate = CStr(Format(Date + 1, "dd/mm/yyyy")) 'help user with default data entry!
    End If
    EnDate = wshRep.Range("C" & iRow).Value
    If EnDate = "" Then
        EnDate = CStr(Format(Date + 10, "dd/mm/yyyy")) 'help user with default data entry!
    End If
    Owner = wshRep.Range("D" & iRow).Value
    Activity = wshRep.Range("E" & iRow).Value
    Comments = wshRep.Range("F" & iRow).Value
    Status = wshRep.Range("G" & iRow).Value
    Situation = wshRep.Range("H" & iRow).Value
    STGoal = wshRep.Range("I" & iRow).Value
    LTGoal = wshRep.Range("J" & iRow).Value
    Expenses = wshRep.Range("K" & iRow).Value
    HoursSpend = wshRep.Range("L" & iRow).Value
    ValueAdded = wshRep.Range("M" & iRow).Value
    picPath = wshRep.Range("N" & iRow).Value

ElseIf SheetName = "Report" And isMatchingIdOnPlanning = False Then

    ' User attempt to create new action from WS Report
    StDate = wshRep.Range("B" & iRow).Value
    If StDate = "" Then
        StDate = CStr(Format(Date + 1, "dd/mm/yyyy")) 'help user with default data entry!
    End If
    EnDate = wshRep.Range("C" & iRow).Value
    If EnDate = "" Then
        EnDate = CStr(Format(Date + 10, "dd/mm/yyyy")) 'help user with default data entry!
    End If
    Owner = wshRep.Range("D" & iRow).Value
    Activity = wshRep.Range("E" & iRow).Value
    Comments = wshRep.Range("F" & iRow).Value
    Status = wshRep.Range("G" & iRow).Value
    Situation = wshRep.Range("H" & iRow).Value
    STGoal = wshRep.Range("I" & iRow).Value
    LTGoal = wshRep.Range("J" & iRow).Value
    Expenses = wshRep.Range("K" & iRow).Value
    HoursSpend = wshRep.Range("L" & iRow).Value
    ValueAdded = wshRep.Range("M" & iRow).Value
    picPath = wshRep.Range("N" & iRow).Value

Else
    
    MsgBox "What Else?!?!?!"
    Item = ConvertString(Item)
    

End If
'
' --------------------
' Populate User Form with values from Variables
PlanningForm.tboxItem.Text = Item                    ' Item
PlanningForm.tboxRow.Value = iRow                    ' Shows actual Row of the data
PlanningForm.tboxSheet.Value = SheetName             ' Workheet Name
PlanningForm.cboxOwner.Text = Owner                  ' Activity Owner
PlanningForm.tboxSTGoal.Text = STGoal                ' Text with Short Term Goal
PlanningForm.cboxLTGoal.Text = LTGoal                ' Text with Long Term Goal
PlanningForm.cboxStatus.Value = Status               ' Status in ComboBox
PlanningForm.cboxSituation.Value = Situation         ' Priority in ComboBox
PlanningForm.tboxActivity.Value = Activity           ' Text with Activity
PlanningForm.tboxComments.Text = Comments            ' Text with Comment
PlanningForm.tboxStartDate.Value = StDate            ' Start Date of activity
PlanningForm.tboxEndDate.Value = EnDate              ' End Date of activity
PlanningForm.tboxExpense.Value = Expenses            ' Expenses in $ terms
PlanningForm.tboxHrsSpend.Value = HoursSpend         ' Time spend in Hours
PlanningForm.tboxValueAdd.Value = ValueAdded         ' Value added in $ terms
PlanningForm.tboxPath.Value = picPath                ' String with picture path [not shown to the user]
' --------------------
' Bring default picture
If Not picPath = "" Then
PlanningForm.imageReport.Picture = LoadPicture(picPath)
Else
PlanningForm.imageReport.Picture = LoadPicture([Summary!J2].Value)
End If
' --------------------
' Put color index number and background colors
PlanningForm.tboxStatusKey.Value = getStatusKey(Status)
PlanningForm.tboxSituationKey.Value = getStatusKey(Situation)
PlanningForm.tboxStatusKey.BackColor = getvbColor(Status)
PlanningForm.tboxSituationKey.BackColor = getvbColor(Situation)
' --------------------
' Populate User Form controls by guessing values:
' --------------------
' Returning a Archive Option
If Status = "Canceled" Then
PlanningForm.optionYes.Value = True
Else
PlanningForm.optionYes.Value = False
End If
                 
If Not Status = "Canceled" Then
PlanningForm.optionNo.Value = True
Else
PlanningForm.optionNo.Value = False
End If

End Sub

'========================================
'CREATE POWER POINT SLIDES
'========================================

' This Sub creates PowerPoint slide from the given row (iRow) of the Report page
' Important: Go to Tools -> References -> Enable Microsoft PPT Object Library

Sub WorkbooktoPowerPoint(iRow)
    
' Declare variables
    'for PowerPoint slides
    Dim PPT As Object: Dim PPTPres As Object: Dim PPTSlide As Object
    Dim oPicture As PowerPoint.Shape: Dim tboxComment As PowerPoint.Shape
    Dim tboxStatus As PowerPoint.Shape: Dim tboxPriority As PowerPoint.Shape
    Dim tboxCost As PowerPoint.Shape: Dim figCircle1 As PowerPoint.Shape: Dim figCircle2 As PowerPoint.Shape
    Dim Key As Integer
    ' for Excel worksheet
    Set wb = ThisWorkbook
    Set wshRep = wb.Worksheets("Report")
    Dim Item As String: Dim Owner As String: Dim Issue As String: Dim Comment As String: Dim Cost As String
    Dim Status As String: Dim Priority As String
    Dim picPath As String
    
' Open PowerPoint, Add Presentation, Make it visible
    Set PPT = CreateObject("PowerPoint.Application")
    Set PPTPres = PPT.Presentations.Add
    PPT.Visible = True
        
' Set the data to variables
    Item = wshRep.Range("A" & iRow).Value
    Owner = wshRep.Range("D" & iRow).Value
    Key = wshRep.Range("G" & iRow).Interior.ColorIndex ' save color property value to Key variable 'CASE STUDY 3
    Issue = wshRep.Range("E" & iRow).Value
    Comment = wshRep.Range("F" & iRow).Value
    Status = wshRep.Range("G" & iRow).Value
    Priority = wshRep.Range("H" & iRow).Value
    Cost = wshRep.Range("K" & iRow).Value
    picPath = wshRep.Range("N" & iRow).Value ' extend this list
    
' Add new blank slide and set the title
    Set PPTSlide = PPTPres.Slides.Add(Index:=1, Layout:=ppLayoutTitleOnly)
    PPTSlide.Select: PPTSlide.Shapes.Title.TextFrame.TextRange.Text = Owner & " - item ID: " & Item & " - " & Left(Issue, 50) & "..."
         
' Paste the picture and adjust its position
On Error Resume Next
    Set oPicture = PPTSlide.Shapes.AddPicture(picPath, msoFalse, msoTrue, Left:=100, Top:=150, Width:=400, Height:=300)

' Add text box for Comment
    Set tboxComment = PPTSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=500, Top:=150, Width:=400, Height:=250)
    
    ' Format the text range
    With tboxComment.TextFrame.TextRange
        .Text = "Comment: " & Left(Comment, 90) & "..."
        With .Font
            .Size = 24
            .Name = "Arial"
        End With
    End With

' Add text box for cost
Set tboxCost = PPTSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=500, Top:=450, Width:=400, Height:=250)

    With tboxCost.TextFrame.TextRange
        .Text = "Approx.cost: " & Cost & " CHF"
        With .Font
            .Size = 24
            .Name = "Arial"
        End With
    End With
    
' Add text box for cost
Set tboxStatus = PPTSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=550, Top:=300, Width:=100, Height:=25)

    With tboxStatus.TextFrame.TextRange
        .Text = "Status: " & Status
        With .Font
            .Size = 12
            .Name = "Arial"
        End With
    End With
    
' Add circle with issue color code
Set figCircle1 = PPTSlide.Shapes.AddShape(Type:=msoShapeOval, Left:=550, Top:=350, Width:=70, Height:=70)
          'Decide which color
          figCircle1.Fill.ForeColor.RGB = getvbColor(Status)

' Add text box for cost
Set tboxPriority = PPTSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=750, Top:=300, Width:=100, Height:=25)

    With tboxPriority.TextFrame.TextRange
        .Text = "Priority: " & Priority
        With .Font
            .Size = 12
            .Name = "Arial"
        End With
    End With
' Add circle with issue color code
Set figCircle2 = PPTSlide.Shapes.AddShape(Type:=msoShapeOval, Left:=750, Top:=350, Width:=70, Height:=70)
          'Decide which color
          figCircle2.Fill.ForeColor.RGB = getvbColor(Priority)

'Step 5.4: Apply Template
On Error Resume Next
' set your path...
PPTPres.Application.ActivePresentation.ApplyTemplate "C:\Users\fxtrams\Downloads\WidescreenPresentation.potx"

' Memory Cleanup (Step useful when adding for loop)
    PPT.Activate
    Set PPTSlide = Nothing
    Set PPTPres = Nothing
    Set PPT = Nothing
               
End Sub

'========================================
'CREATE PIVOT TABLE
'========================================

' This Sub creates Pivot Table to be used for visualization of the activities
'
Sub CreatePivotTable()

Dim lRow As Long, lCol As Long
Dim TableRange As Range
Dim PTNumber As Integer
Dim PTRNumber As Integer
Set wb = ThisWorkbook
Set wshRep = wb.Worksheets("Report")
Set wshPiv = wb.Worksheets("Pivot")
Dim PC As PivotCache
Dim PT As PivotTable

' Setup the Table Number
PTNumber = wshPiv.CELLS(Rows.Count, 1).End(xlUp).Value
PTNumber = PTNumber + 1
PTRNumber = PTNumber * 10

'Store our data in Range variable:
With wshRep
    lRow = .CELLS(Rows.Count, "A").End(xlUp).Row
    lCol = .CELLS(1, Columns.Count).End(xlToLeft).Column
    Set TableRange = .Range(.CELLS(1, 1), .CELLS(lRow, lCol))
End With

' Set Pivot Table Cache
Set PC = wb.PivotCaches.Create(xlDatabase, TableRange)

' Set Pivot Table
Set PT = PC.CreatePivotTable(wshPiv.Range("B" & PTRNumber), "MyGoals" & PTNumber)

' Write PTNumber to the worksheet
wshPiv.CELLS(PTNumber, 1).Value = PTNumber

'// Adding Columns, Rows and Data to pivot table
With PT

    '// Pivot Table Layout
    .RowAxisLayout xlTabularRow
    .ColumnGrand = False 'Optional (Column Grand Total)
    .RowGrand = False 'Optional (Row Grand Total)

    .TableStyle2 = "PivotStyleMedium9"
    .HasAutoFormat = False 'Re-Format Pivot Table when refresh
    .SubtotalLocation xlAtTop 'Position SubTotal on the top or bottom

End With


ClearObjects:
Set PC = Nothing
Set PT = Nothing
Set TableRange = Nothing

Call clear_objects

Exit Sub

errHandle:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo ClearObjects

End Sub

Private Sub clear_objects()

' Release Memory
Set wshRep = Nothing
Set wb = Nothing

End Sub
