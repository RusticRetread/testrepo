Attribute VB_Name = "Module1"
Option Compare Text
Private arrCourses() As objCourse
Private arrCompulsoryCourses() As objCourse
Private GenderWeight As Double
Private PlacesWeight As Double
Private TopScore As Double
Private TopIndex(2) As Integer
Const clcColRank As Integer = 3
Const clcColUsername As Integer = 5
Const clcColClass As Integer = 4
Const clcColFName As Integer = 2
Const clcColSName As Integer = 1
Const clcResultsStart As Integer = 13
Const clcCopyStart As Integer = 6
Const simsColClass As Integer = 4
Const simsColFName As Integer = 2
Const simsColSName As Integer = 1
Const simsPasteStart As Integer = 10
Const simsColGender As Integer = 3
Const simsAlevelsStart As Integer = 5
Const simsColRank As Integer = 18
Const simsResultsStart As Integer = 20
Const dispQtrOneRow As Integer = 4
Const dispQtrSpacing As Integer = 30
Const dispStartCell As String = "a2"



Sub ChangeNameFormat()
Dim NameArray() As String
Dim FirstNameString As String
Dim arrUbound As Integer
Dim i As Integer
Dim rowCount As Integer
rowCount = 2
Dim clcRange As Range

Set clcRange = clcsheet.Cells
clcRange.Interior.ColorIndex = xlColorIndexNone

'checks to see if this has been run before - looks for surname at begining
'if not then assumes not run and inserts two columns at beginnig
If Not clcsheet.Cells(1, 1) = "Surname" Then
    clcsheet.Range("A:B").EntireColumn.Insert
Else
    clcsheet.Range("A:B").EntireColumn.Clear
End If
clcsheet.Range("A1").Value = "Surname"
clcsheet.Range("B1").Value = "Firstname"
'clcsheet.Range("e1").Value = "Class"

Do Until clcsheet.Cells(rowCount, clcColUsername) = ""
    NameArray() = Split(clcsheet.Cells(rowCount, clcColUsername).Value, " ")
    arrUbound = UBound(NameArray)
    clcsheet.Cells(rowCount, 1).Value = NameArray(arrUbound)
    For i = 0 To arrUbound - 1
        If i = 0 Then
            FirstNameString = NameArray(i)
        Else
            FirstNameString = FirstNameString & " " & NameArray(i)
        End If
    Next
    clcsheet.Cells(rowCount, 2).Value = FirstNameString
    
    

rowCount = rowCount + 1
Loop
    

End Sub

Function Tidy(p_string) As String
Tidy = Trim(UCase(p_string))
End Function
Sub CLCtoSIMS()
Dim SimsRowCount As Integer
Dim CLCrowCount As Integer
Dim SimsFName As String
Dim SimsSName As String
Dim ClcFname As String
Dim ClcSname As String
Dim ClcClass As String
Dim SimsClass As String
Dim CollRowIndex As Variant
Dim copyRow As Integer
Dim copyVal As String
Dim aLevel As String
Dim UsedNameIndex As Variant
Dim ClcNameIndexUsed As Boolean
Dim x, y As Integer
Dim clcRange As Range





Set clcRange = clcsheet.Cells
clcRange.Interior.ColorIndex = xlColorIndexNone

With SimsSheet
    .Range(.Cells(2, simsPasteStart), .Cells(10000, 50)).Clear
    .Range(.Cells(2, 1), .Cells(10000, 50)).Interior.ColorIndex = xlColorIndexNone
End With

copyRow = -1

Dim NameColl As New Collection
Dim UsedNamesIndexColl As New Collection
Dim DuplicateSubjects As New Collection



'If Not clcsheet.Cells("C3").Value = "Surname" Then
    'MsgBox "you forgot to do the name swap"
    'Exit Sub
'End If

CLCrowCount = 2
SimsRowCount = 2
ClcNameIndexUsed = False

Do Until SimsSheet.Cells(SimsRowCount, 1) = ""
    SimsSName = Tidy(SimsSheet.Cells(SimsRowCount, simsColSName))
    SimsFName = Tidy(SimsSheet.Cells(SimsRowCount, simsColFName))
    SimsClass = Tidy(SimsSheet.Cells(SimsRowCount, simsColClass))
    
        Do Until clcsheet.Cells(CLCrowCount, 3) = ""
            ClcSname = Tidy(clcsheet.Cells(CLCrowCount, clcColSName))
            ClcFname = Tidy(clcsheet.Cells(CLCrowCount, clcColFName))
            ClcClass = Tidy(clcsheet.Cells(CLCrowCount, clcColClass))
            If ClcSname = SimsSName And ClcFname = SimsFName And ClcClass = SimsClass Then
                NameColl.Add (CLCrowCount) 'store the row of the hit
            End If
            CLCrowCount = CLCrowCount + 1
        Loop
        
        CLCrowCount = 2
        
        If NameColl.Count = 0 Then
            'nothing found so make a note
            SimsSheet.Cells(SimsRowCount, simsPasteStart).Value = "No Match"
            Else
                    If NameColl.Count > 1 Then
                         'we have found duplicate names in clc for the entry in sims so mark the clc entries with a colour
                         'and note the duplicate in sims
                         For Each CollRowIndex In NameColl
                                With clcsheet.Rows(CollRowIndex).Interior
                                    .Color = RGB(200, 20, 20)
                                    .TintAndShade = 0.5
                                 End With
                                 SimsSheet.Cells(SimsRowCount, simsPasteStart).Value = "DUPLICATE IN CLC"
                                 Debug.Print "clc duplicate in row " & CollRowIndex & " found for " _
                                 & SimsSheet.Cells(SimsRowCount, simsColSName) & " at sims row " & SimsRowCount
                         Next
                     Else
                        'now we can only have one row number in the collection but still have to check against duplicates in sims - check a collection to see if this name in clc has
                        'been used before (via row number) - if it has then a previous identical combination of name and class
                        'from sims has hit this entry in clc - must mean a duplicate in sims
                        copyRow = NameColl(1)
                         For Each UsedNameIndex In UsedNamesIndexColl
                            If copyRow = UsedNameIndex Then
                                ClcNameIndexUsed = True
                            End If
                         Next
                         If ClcNameIndexUsed = True Then
                              clcsheet.Rows(copyRow).Interior.Color = RGB(200, 0, 0)
                              SimsSheet.Rows(SimsRowCount).Interior.Color = RGB(200, 0, 0)
                              SimsSheet.Cells(SimsRowCount, simsPasteStart).Value = "2ND DUPLICATE IN SIMS"
                              Debug.Print " Clc Name in row " & copyRow & " was matched with " _
                              & SimsSheet.Cells(SimsRowCount, simsColSName) & " at sims row " & SimsRowCount & _
                              " but has been already used"
                              ClcNameIndexUsed = False
                                 
                         Else
                                'copy choices accross
                                y = 0
                                For x = clcCopyStart To clcCopyStart + 7
                                    copyVal = Tidy(clcsheet.Cells(copyRow, x).Value)
                                    'exclude duplicates and spaces when copying choices accross
                                    If Not Len(copyVal) = 0 Then
                                        If Not DuplicateExists(copyVal, DuplicateSubjects) Then
                                            SimsSheet.Cells(SimsRowCount, simsPasteStart + y).Value = copyVal
                                            DuplicateSubjects.Add copyVal, copyVal
                                            y = y + 1
                                        End If
                                    End If
                                Next
                                UsedNamesIndexColl.Add (copyRow)
                                SimsSheet.Cells(SimsRowCount, simsColRank).Value = clcsheet.Cells(copyRow, clcColRank).Value
                                clcsheet.Cells(copyRow, 1).Interior.Color = RGB(152, 251, 152)
                                Set DuplicateSubjects = New Collection
                         End If
                           
                     End If
        End If
        
        
        
        Set NameColl = New Collection
        copyRow = -1
        SimsRowCount = SimsRowCount + 1
Loop
'sort worksheet
With SimsSheet
.Range(.Cells(2, 1), .Cells(SimsRowCount - 1, simsColRank)).Sort (.Cells(2, simsColRank))
End With
        
'remember to sort simssheet
 
End Sub
Function DuplicateExists(p_subjectChoice As String, ByRef p_Coll As Collection) As Boolean
On Error GoTo EH
If Len(p_Coll.Item(p_subjectChoice)) > 0 Then
    DuplicateExists = True
End If
Exit Function
EH:
DuplicateExists = False
End Function



Function ExcludeChoice(ByRef p_StudentObj As objStudent, p_cChoice As String, p_SimsRow As Integer) As Boolean
Dim x As Integer
ExcludeChoice = False
If p_StudentObj.ExcludedCombo(p_cChoice) Then
    For x = simsPasteStart To simsPasteStart + 7
        If Tidy(SimsSheet.Cells(p_SimsRow, x).Value) = p_cChoice Then
            SimsSheet.Cells(p_SimsRow, x).Interior.Pattern = xlPatternCrissCross
        End If
    Next
    ExcludeChoice = True
End If

End Function
Sub SetUpExclusions(p_Stud As objStudent, p_Coll As Collection, p_rowCount As Integer)
Dim subChoice As String
On Error Resume Next
Dim x As Integer
 For x = simsAlevelsStart To simsAlevelsStart + 4
        subChoice = "EMPTY"
        subChoice = p_Coll.Item(SimsSheet.Cells(p_rowCount, x).Value)
        If Not subChoice = "EMPTY" Then
            p_Stud.AddExcludedCombo (subChoice)
        End If
 Next
End Sub

Sub Main()
'
Dim ArrName As String
Dim iCount As Integer
'Dim AllocRowCount As Integer
Dim AllocColCount As Integer
Dim sCourseName As String
Dim Course As Variant
Dim varCourse As Variant
Dim arrCrsTrk(3, 4) As Variant
Dim Student As objStudent
Dim StudentColl As New Collection
Dim CourseNames As New Collection
Dim ExcludedSubjects As New Collection
GenderWeight = Range("GenderWeight").Value
Dim i As Integer
Dim x As Integer
Dim y As Integer

'find out what size to make the elective and compulsory course arrays and then fill them
'Set Student = New objStudent
'testsub Student
y = 0
x = 1
Dim RangeStart As Range
Dim CrsRange As Range
Dim RangeCell As Variant
Set RangeStart = ctrlsheet.Range("A1").Offset(1, 0)
Set CrsRange = Range(RangeStart, RangeStart.End(xlDown))
For Each RangeCell In CrsRange
    DispSheet.Cells(2, x).Value = Tidy(RangeCell.Value)
    x = x + 1
    If RangeCell.Offset(0, 2) = "Y" Then
        y = y + 1
    End If
Next

ReDim Preserve arrCourses(1 To (CrsRange.Count - y) * 4)
ReDim Preserve arrCompulsoryCourses(1 To y * 4)
    

y = 1
x = 1
For Each RangeCell In CrsRange
        For i = 1 To 4
            Set Course = New objCourse
            
            Course.Name = Tidy(RangeCell.Value)
            Course.Quarter = i
            Course.TotalPlaces = RangeCell.Offset(0, 1)
            Course.SetInitialScore RangeCell.Offset(0, 1)
            If RangeCell.Offset(0, 2) = "Y" Then
                Course.Index = x
                Set arrCompulsoryCourses(x) = Course
                x = x + 1
            Else
                Course.Index = y
                Set arrCourses(y) = Course
                y = y + 1
            End If
        Next
    Next
Set Course = Nothing

iCount = 3
Do Until ctrlsheet.Cells(iCount, 4).Value = ""
    'add the a level choice to the collection with the carousel choice as a key
    ExcludedSubjects.Add Tidy(ctrlsheet.Cells(iCount, 5)), Tidy(ctrlsheet.Cells(iCount, 4))
    iCount = iCount + 1
Loop

iCount = 2

Do Until SimsSheet.Cells(iCount, simsColSName) = ""
    'simspastestart is where the clc repsonses were copied to
    'If Tidy(simssheet.Cells(iCount, simsPasteStart)) = "NO RESPONSE" Then
        'put them on scrap pile
    'Else
        Set Student = New objStudent
        'To Do - check that rank number agrees with position in spreadsheet
    
            
            Dim iRank As Integer
            iRank = SimsSheet.Cells(iCount, simsPasteStart + 8)
            If Not iCount - 1 = iRank Then
                MsgBox "students aren't ranked at " & iRank
                Exit Sub
            End If
                        
            With Student
                .Gender = Tidy(SimsSheet.Cells(iCount, simsColGender))
                .Index = iCount - 1 'should be the same as the row number
                .FirstName = SimsSheet.Cells(iCount, simsColFName)
                .Surname = SimsSheet.Cells(iCount, simsColSName)
                .Rank = iRank
                .Class = SimsSheet.Cells(iCount, simsColClass)
            End With
            
            SetUpExclusions Student, ExcludedSubjects, iCount
            
            
            x = 0
            y = 0
            Do Until x > 6
                        'rely on sorting spreadsheet to rank students but validate this
                        'simssheet should have subject choices without duplicates or spaces inbetween choices but still with excluded courses
                        If y < 7 Then
                            sCourseName = Tidy(SimsSheet.Cells(iCount, simsPasteStart + y))
                            If Len(sCourseName) = 0 Then sCourseName = "NULL"
                        Else
                            sCourseName = "NULL"
                        End If
                        If Not ExcludeChoice(Student, sCourseName, iCount) Then
                            Student.SetCrsChoice sCourseName, x
                            x = x + 1
                        
                        End If
                        y = y + 1 'always increment counter that searches through choices but only increment the array in student choices when there is something to write (inc null)
             Loop
            
            Debug.Print "Student Id *********************" & Student.Index
            StudentColl.Add Student
            'clean up the score indexes
            TopScore = 0
            'call the recursive sub that will populate student obj sel courses index with the indices of the best three course
            'combination using topscore to sort the course scores. the empty tracking array is passed in to allow later recursion
            If Tidy(Student.GetCrsChoice(0)) = "NO RESPONSE" Then
                'recursive  search - prev worksheet sort should mean these only happen right at end so hopefully not too long
                fullRecurse arrCrsTrk, 0, 0, Student, True
            Else
                 fullRecurse arrCrsTrk, 0, 0, Student, False
            End If
            'possible rerun using next choice as primary if only 2 courses selected
            'run non recursive search for best fit interview course and add to topindex array
            
        y = 0
        For i = 0 To 2
            If Student.GetSelCourse(i) > 0 Then
                Set Course = arrCourses(Student.GetSelCourse(i))
                Course.addStudent Student.Index, Student.Gender, GenderWeight
                With SimsSheet.Cells(iCount, (simsResultsStart - 1) + Course.Quarter)
                .Value = Course.Name & " " & Course.Quarter & " " & Course.Index
                '.Interior.ColorIndex = Course.Quarter + 2
                End With
                             
                For y = simsPasteStart To simsPasteStart + 7
                    If SimsSheet.Cells(iCount, y) = Course.Name Then SimsSheet.Cells(iCount, y).Interior.ColorIndex = (i + 3)
                Next
                                'write chosen course to appropriate cycle in sims sheet  and clc as crosscheck and indicate chosen course with index based color on prefs
                ' working out if we have any students without a choice - we don't add the chosen courses to the student obj - should we?
            End If
        Next
            Set Course = arrCompulsoryCourses(Student.CompulsoryCourseIndex)
            Course.addStudent Student.Index, Student.Gender, GenderWeight
            SimsSheet.Cells(iCount, (simsResultsStart - 1) + Course.Quarter) = Course.Name & " " & Course.Quarter & " " & Course.Index
            Set Course = Nothing
            If Student.TotalSelCourses < 3 Then
                MsgBox Student.FirstName & " " & Student.Surname & " didn't get three electives"
                'write their names and no of missed courses somewhere?
            End If
            If Student.CompulsoryCourseIndex = 0 Then
                MsgBox Student.FirstName & " " & Student.Surname & " not assigned a compulsory course"
            End If
            'y = 0
            
    'End If
iCount = iCount + 1
Loop
'Dim colorRange As Range
'Set colorRange = DispSheet.Range("a2", "z10000")
'colorRange.Select
'colorRange.Interior.ColorIndex = xlColorIndexNone


DisplayCourses arrCourses(), StudentColl
DisplayCourses arrCompulsoryCourses(), StudentColl
 
End Sub
Sub DisplayCourses(CrsArray As Variant, p_StudentColl As Variant)
Dim RangeStart As Range
Dim CrsRange As Range
Dim Course As Variant
Dim RangeCell As Variant
Dim RowVal As Integer
Dim i As Integer
Dim m_Student As objStudent
Dim StudentIndex As Integer


Set RangeStart = DispSheet.Range(dispStartCell)
Set CrsRange = Range(RangeStart, RangeStart.End(xlToRight))
For Each Course In CrsArray
    For Each RangeCell In CrsRange
        If Tidy(RangeCell.Value) = Course.Name Then
            Select Case Course.Quarter
                Case 1
                RowVal = dispQtrOneRow
                Case 2
                RowVal = dispQtrOneRow + dispQtrSpacing
                Case 3
                RowVal = dispQtrOneRow + dispQtrSpacing * 2
                Case 4
                RowVal = dispQtrOneRow + dispQtrSpacing * 3
            End Select
            
            If DispSheet.Cells(RowVal, RangeCell.Column).Value = "" Then
                For i = 0 To (Course.TotalPlaces - 1) 'course student array starts at 0
                    StudentIndex = Course.getStudentIndex(i)
                    If StudentIndex > 0 Then
                    Set m_Student = p_StudentColl(Course.getStudentIndex(i))
                        If m_Student.Index = Course.getStudentIndex(i) Then 'make sure collection position = student index
                            DispSheet.Cells(RowVal, RangeCell.Column).Value = m_Student.Surname & " " _
                                & m_Student.FirstName & " " & m_Student.Class ' & " " & m_Student.Index
                        Else
                            MsgBox "student collection position and index disagree for " & m_Student.Index
                        End If
                    Else
                        
                        Set colorRange = DispSheet.Cells(RowVal, RangeCell.Column)
                        colorRange.Interior.ColorIndex = 15
                        colorRange.Value = "***********"
                        
                    End If
                    RowVal = RowVal + 1
                 Next
                 DispSheet.Cells(RowVal, RangeCell.Column).Value = "Gender Bal " & Course.GenderBalance & " Slots Rem " & Course.RemainingSlots
                 Exit For
            End If
            
        End If
    Next
Next

End Sub


Sub fullRecurse(valsArray() As Variant, targetIndex As Integer, arrayIndex As Integer, p_StudentObj As objStudent, AllCourses As Boolean)

Dim iCount As Integer
Dim newtarget As Integer
Dim targetString As String
Dim valuefound As Boolean
Dim crsObj As Variant
Dim Score As Double
Dim CompulsoryCrsIndex As Integer
Dim i As Integer


'CONSIDER SPLITTING COURSES IN TO 4 ARRAYS BY QUARTER TO CUT DOWN SEARCHING WHOLE
'ARRAY STACK EACH TIME!!!!



iCount = 1
If AllCourses = False Then
    targetString = Tidy(p_StudentObj.GetCrsChoice(targetIndex)) ' should be something there or nulls
    'If targetString = "" Then targetString = "ALL"
Else
    targetString = "NULL"
End If

For Each crsObj In arrCourses

If ValidCourse(Tidy(crsObj.Name), crsObj.IsFull, targetString, p_StudentObj) Then ' consider putting this in a boolean function to allow eventual use of same module for "no response"
    If ValidQuarter(crsObj.Quarter, valsArray(), crsObj.Name) Then ' this should stay as a separate step to avoid running it for every course object in array
                                                       'also checks valsarray to see if we have already selected another instance of this course
    valuefound = True
    'Debug.Print arrayIndex
    valsArray(arrayIndex, 0) = crsObj.Name
    valsArray(arrayIndex, 1) = crsObj.Index
    valsArray(arrayIndex, 2) = crsObj.Quarter
    valsArray(arrayIndex, 3) = crsObj.GetScore(p_StudentObj.Gender)
    
    If targetIndex = 6 Or (arrayIndex + 1) = 3 Then 'maybe call a "all courses" recurse if run out of preferences and still options empty
        Score = valsArray(0, 3) + valsArray(1, 3) + valsArray(2, 3)
        Debug.Print valsArray(0, 1) & " " & valsArray(1, 1) & " " & valsArray(2, 1)
        Debug.Print Score
        
        If Score > TopScore Or TopScore = 0 Then
            CompulsoryCrsIndex = ValidCompulsoryCourse(valsArray(), p_StudentObj.Gender)
            If Not CompulsoryCrsIndex = 0 Then
                    TopScore = Score
                    For i = 0 To 2
                        If IsNumeric(valsArray(i, 1)) Then
                            p_StudentObj.SetSelCourse i, Int(valsArray(i, 1))
                            'p_StudentObj.SetSelCourse 1, 0
                        Else
                            p_StudentObj.SetSelCourse i, 0
                        End If
                    Next
                    p_StudentObj.CompulsoryCourseIndex = CompulsoryCrsIndex
                    Debug.Print p_StudentObj.Index
                    Debug.Print iCount & " " & arrayIndex
                    Debug.Print valsArray(0, 0) & " " & valsArray(1, 0) & " " & valsArray(2, 0)
                    Debug.Print valsArray(0, 2) & " " & valsArray(1, 2) & " " & valsArray(2, 2)
                    Debug.Print valsArray(0, 1) & " " & valsArray(1, 1) & " " & valsArray(2, 1)
                    Debug.Print valsArray(0, 3) & " " & valsArray(1, 3) & " " & valsArray(2, 3)
                    Debug.Print Score
                    Debug.Print "-------------------------------------------------"
            End If
        End If
    Else
        fullRecurse valsArray(), targetIndex + 1, arrayIndex + 1, p_StudentObj, AllCourses
    End If
    
    valsArray(arrayIndex, 0) = ""
    valsArray(arrayIndex, 1) = -1
    valsArray(arrayIndex, 2) = 0
    valsArray(arrayIndex, 3) = 0
    End If
End If

iCount = iCount + 1
Next

If valuefound = False Then
    If targetIndex = 6 Or (arrayIndex) = 3 Then 'arrayindex check redundant i think
        Score = valsArray(0, 3) + valsArray(1, 3) + valsArray(2, 3)
        If Score > TopScore Or TopScore = 0 Then
            CompulsoryCrsIndex = ValidCompulsoryCourse(valsArray(), p_StudentObj.Gender)
            If Not CompulsoryCrsIndex = 0 Then
                    TopScore = Score
                    For i = 0 To 2
                        If IsNumeric(valsArray(i, 1)) Then
                            p_StudentObj.SetSelCourse i, Int(valsArray(i, 1))
                            'p_StudentObj.SetSelCourse 1, 0
                        Else
                            p_StudentObj.SetSelCourse i, 0
                        End If
                    Next
                    p_StudentObj.CompulsoryCourseIndex = CompulsoryCrsIndex
                    Debug.Print p_StudentObj.Index
                    Debug.Print iCount & " " & arrayIndex
                    Debug.Print valsArray(0, 0) & " " & valsArray(1, 0) & " " & valsArray(2, 0)
                    Debug.Print valsArray(0, 2) & " " & valsArray(1, 2) & " " & valsArray(2, 2)
                    Debug.Print valsArray(0, 1) & " " & valsArray(1, 1) & " " & valsArray(2, 1)
                    Debug.Print valsArray(0, 3) & " " & valsArray(1, 3) & " " & valsArray(2, 3)
                    Debug.Print Score
                    Debug.Print "-------------------------------------------------"
            End If
        End If
    Else
    fullRecurse valsArray, targetIndex + 1, arrayIndex, p_StudentObj, AllCourses
    End If
End If

End Sub

Function ValidQuarter(qtr As Integer, ByRef chkarray As Variant, Optional p_crsName As Variant) As Boolean
Dim i As Integer
ValidQuarter = True

    For i = LBound(chkarray) To UBound(chkarray)
        If IsMissing(p_crsName) Then
            If qtr = chkarray(i, 2) Then
                ValidQuarter = False
            End If
        Else
            If qtr = chkarray(i, 2) Or p_crsName = chkarray(i, 0) Then
                ValidQuarter = False
            End If
        End If
    Next
    
End Function

Function ValidCourse(crsName As String, crsFull As Boolean, TgtString As String, p_StudObj As objStudent) As Boolean
ValidCourse = False
If Not p_StudObj.ExcludedCombo(crsName) Then
    If TgtString = "NULL" Then
        If crsFull = False Then ValidCourse = True
    Else
        If crsFull = False And crsName = TgtString Then ValidCourse = True
    End If
End If
End Function

Function ValidCompulsoryCourse(p_chkArray As Variant, p_Gender As String)
Dim objCompulsoryCrs As Variant
Dim TopScore As Double
Dim TopIndex As Integer
Dim RunningScore As Double
TopScore = 0
TopIndex = 0
For Each objCompulsoryCrs In arrCompulsoryCourses
    If Not objCompulsoryCrs.IsFull And ValidQuarter(objCompulsoryCrs.Quarter, p_chkArray) Then
        RunningScore = objCompulsoryCrs.GetScore(p_Gender)
        If RunningScore > TopScore Or TopScore = 0 Then
            TopIndex = objCompulsoryCrs.Index
            TopScore = RunningScore
        End If
    End If
Next
ValidCompulsoryCourse = TopIndex
'course indices start at 1 so 0 will indicate that no valid course found
End Function

Sub FinalValidation()
Dim SearchRange As Range
Dim rowCount As Integer
Dim SimsRowCount As Integer
Dim SearchString As String
Dim RangeStart As Range
Dim RangeEndCol As Integer
Dim RangeStartRow As Integer
Dim RangeEndRow As Integer
Dim RangeColl As New Collection
Dim rangeResult As Range
Dim CourseName As String
Dim SimsNameAlreadyFound As Boolean
Dim i As Integer
Dim y As Integer
Dim z As Integer
Dim FirstAddress As Variant
'Dim c As Range
'Dim firstaddress As Variant
'
'
''Set hatevba = DispSheet.Range("A4", "E22")
'' Set SearchRange = hatevba.Find("blah", searchorder:=xlByColumns)
'' Debug.Print SearchRange.Address
''Set SearchRange = hatevba.FindNext(SearchRange)
''Debug.Print SearchRange.Address
'' MsgBox Trim("AAA ") = "aaa"
'With DispSheet.Range("a22:b32")
'     Set c = .Find(2, LookIn:=xlValues)
'     If Not c Is Nothing Then
'        firstaddress = c.Address
'        Do
'            c.Value = 5
'            Set c = .FindNext(c)
'        If c Is Nothing Then
'            GoTo DoneFinding
'        End If
'        Loop While c.Address <> firstaddress
'      End If
'DoneFinding:
'End With





Set RangeStart = DispSheet.Range(dispStartCell) 'look at the first cell of subject headers
RangeEndCol = RangeStart.End(xlToRight).Column 'find last col of subject headers
rowCount = 2

Do Until clcsheet.Cells(rowCount, 1) = "" 'loop through clcsheet building searchstring for each entry
                SearchString = Trim(clcsheet.Cells(rowCount, clcColSName)) & " " & Trim(clcsheet.Cells(rowCount, clcColFName)) & " " & Trim(clcsheet.Cells(rowCount, clcColClass))
                For i = 0 To 3 'loop through subject quarters using multiples of first quarter row and quarter spacing
                    With DispSheet
                        RangeStartRow = dispQtrOneRow + (dispQtrSpacing * i)
                        RangeEndRow = (dispQtrOneRow + (dispQtrSpacing * (i + 1))) - 1
                        Set SearchRange = Range(.Cells(RangeStartRow, 1), .Cells(RangeEndRow, RangeEndCol))
                    End With
                        
                        Set rangeResult = SearchRange.Find(SearchString, LookAt:=xlPart, searchorder:=xlByColumns)    'look for a name in quarter
                        
                        'FirstAddress = rangeResult.Address
                        If rangeResult Is Nothing Then
                            MsgBox SearchString & " not found in quarter " & i + 1
                        Else
                            'look up the course name and check for duplicate
                            Debug.Print rangeResult.Address
                            CourseName = DispSheet.Cells(RangeStart.Row, rangeResult.Column).Value
                            If Not DuplicateExists(CourseName, RangeColl) Then ' use the duplicateexists function to check if this subject has been already taken
                                  RangeColl.Add CourseName, CourseName
                                  'check for a second instance in the quarter
                                  FirstAddress = rangeResult.Address
                                  Set rangeResult = SearchRange.FindNext(rangeResult)
                                  Debug.Print rangeResult.Address
                                  If Not rangeResult.Address = FirstAddress Then
                                      MsgBox SearchString & " found twice in in quarter " & i + 1
                                  Else
                                      'only get here if there is 1) an instance 2) with a course name not already found 3) no second instance of the name in the quarter
                                      'highlight the choice on clcsheet
                                      For y = clcCopyStart To clcCopyStart + 7
                                              If CourseName = "INTERVIEW SKILLS" Then
                                                  clcsheet.Cells(rowCount, clcCopyStart + 8) = CourseName
                                              Else
                                                    If Tidy(clcsheet.Cells(rowCount, y)) = CourseName Then
                                                        clcsheet.Cells(rowCount, y).Interior.Color = RGB(20, 250, 20)
                                                    End If
                                              'End If
                                          End If
                                      Next
                                                          
                                  End If
                            Else
                                  MsgBox SearchString & " has two instances of the same course"
                            End If
                        End If
                Next
                'we have a collection of courses from disp sheet highlighted on clc sheet. Now x check which students have not got one of their
                'first 4 choices in sims
                SimsRowCount = 2
                Do Until SimsSheet.Cells(SimsRowCount, 1) = ""
                   If SearchString = Trim(SimsSheet.Cells(SimsRowCount, simsColSName)) & " " & Trim(SimsSheet.Cells(SimsRowCount, simsColFName)) _
                   & " " & Trim(SimsSheet.Cells(SimsRowCount, simsColClass)) Then
                        'we know preferences copied to sims have had duplicates excluded so we can just step through choices and make sure that
                        'we get three hits within 4 loops - if not highlight name for manual inspection.  Excluded subjects are marked with background pattern so exclude those.
                        'if we get to 7 loops without 3 hits something has gone wrong so msgbox
                        i = 0
                        z = 0
                        y = 1
                        Do Until i > 6 'i is cell counter
                            If Not SimsSheet.Cells(SimsRowCount, simsPasteStart + i).Interior.Pattern = xlPatternCrissCross Then

                                If DuplicateExists(SimsSheet.Cells(SimsRowCount, simsPasteStart + i).Value, RangeColl) Then
                                    z = z + 1
                                End If
                                If y = 4 And z < 3 Then
                                    SimsSheet.Cells(SimsRowCount, 1).Interior.Color = RGB(255, 192, 0)
                                End If
                                y = y + 1 ' valid preference so increment pref counter

                            End If
                            i = i + 1 'increment cell counter
                        Loop
                    End If
                SimsRowCount = SimsRowCount + 1
                Loop
                
Set RangeColl = New Collection
rowCount = rowCount + 1
Loop
End Sub


