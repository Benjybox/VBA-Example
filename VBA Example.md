# VBA-Example
Here's an example of some VBA code I write, this one is a little old and my current code is a bite more concise, but it's a complex task and quite a few different types of functions, so wanted to show that I can complete complex tasks with my code. This is checking tumour progression status usig the RECIST reporting criteria. 

Dim perIDBL As Variant

Dim NoMoreLines As Integer

Dim RowYN As Integer

Dim expoverallresponse As String

Dim expoverallresponse1st As String

 

Sub RECISTCHECKS()

 

Application.ScreenUpdating = False

 

zoom

saveas

updatedata

checkNumberOfLinesPPt

RemoveMissingUnavailable

dTotals

checkResponse

checkExpTargetRespVeCRF

NTLatBL2

newlesions

NewLesionListed

expectedoverallresponse

CheckExpectedRespAgainsteCRF

wraptext

deletesheets

deleteconnections

deletebutton

 

Application.ScreenUpdating = True

 

MsgBox "ALL FIELDS HAVE BEEN CHECKED, ANYTHING HIGHLIGHTED IN RED NEEDS TO BE REVIEWED BY A DATA MANAGER"

End Sub

 

Sub updatedata()

 

    ActiveWorkbook.RefreshAll

End Sub

Sub zoom()

 

    Sheets("BL RECIST").Select

    ActiveWindow.zoom = 70

    Sheets("12Wk RECIST").Select

    ActiveWindow.zoom = 70

End Sub

Sub saveas()

  Dim dtToday As String

  Dim filename As String

  Dim filepath As String

 

  filepath = "S:\Check 3 RECIST\RECIST Checks " & dtToday

  dtToday = Format(Now(), "DD-MMM-YYYY")

  filename = filepath & dtToday & ".xlsm"

 

ActiveWorkbook.saveas filename

End Sub

 

 

Sub checkNumberOfLinesPPt()

'' this checks that there are 3 lines of data for each patient, this will make sure that the VBA code is easier I think.

 

Worksheets("BL RECIST").Activate

Range("A2").Activate

 

PtNo = 1

count = 0

 

Do

 

    If ActiveCell = PtNo Then

       

        Do

            count = count + 1

            ActiveCell.Offset(1, 0).Activate

   

        Loop Until ActiveCell.Value <> PtNo

   

        If count <> 3 Then

            MsgBox ("Error")

            Exit Sub

           

        End If:

   

    End If:

   

    PtNo = PtNo + 1

    count = 0

   

    If ActiveCell <> PtNo Then

        PtNo = PtNo + 1

    End If:

   

Loop Until ActiveCell = ""

 

Worksheets("12Wk RECIST").Activate

Range("A2").Activate

 

PtNo = 1

count = 0

 

Do

 

    If ActiveCell = PtNo Then

       

        Do

            count = count + 1

            ActiveCell.Offset(1, 0).Activate

   

        Loop Until ActiveCell.Value <> PtNo

   

        If count <> 3 Then

            MsgBox ("Error Patient has <> 3 lines of results")

            Exit Sub

           

        End If:

   

    End If:

   

    PtNo = PtNo + 1

    count = 0

   

    If ActiveCell <> PtNo Then

        PtNo = PtNo + 1

    End If:

   

Loop Until ActiveCell = ""

 

 

 

End Sub

 

Sub RemoveMissingUnavailable()

'

' RemoveMissingUnavailable Macro

'

 

'

 

    Worksheets("BL RECIST").Activate

    Columns("H:J").Select

    Selection.Replace What:="-9", Replacement:="", LookAt:=xlPart, _

        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _

        ReplaceFormat:=False

    Selection.Replace What:="-8", Replacement:="", LookAt:=xlPart, _

        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _

        ReplaceFormat:=False

       

        Worksheets("12Wk RECIST").Activate

    Columns("H:N").Select

    Selection.Replace What:="-9", Replacement:="", LookAt:=xlPart, _

        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _

        ReplaceFormat:=False

    Selection.Replace What:="-8", Replacement:="", LookAt:=xlPart, _

        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _

        ReplaceFormat:=False

       

        

End Sub

 

Sub dTotals()

Dim x As Variant

Dim y As Variant

Dim z As Variant

 

'''''''''' TOTALS FOR BASELINE

 

Worksheets("BL RECIST").Activate

Columns("K:K").ColumnWidth = 15

Columns("K:K").Select

    With Selection.Interior

       .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .ThemeColor = xlThemeColorAccent4

        .TintAndShade = 0.799981688894314

        .PatternTintAndShade = 0

    End With

    With Selection.Font

        .ColorIndex = xlAutomatic

       .TintAndShade = 0

    End With

   

Range("H2").Activate

z = 1

y = 0

 

ActiveCell.Offset(-1, 3).Value = "Total Lesion Measurements at Baseline"

 

 

Do ' start loop

   

    If ActiveCell.Offset(0, -7).Value = "" Then Exit Do

   

    If ActiveCell.Offset(0, -7).Value = z Then ' if pt id = 1 do next

       

        Do

            x = ActiveCell.Value

            y = y + x ' y updated

                       

            ActiveCell.Offset(1, 0).Activate

           

        If ActiveCell.Offset(0, -7).Value = "" Then Exit Do

       

        Loop Until ActiveCell.Offset(0, -7).Value > z ' once person id changes restart

            ActiveCell.Offset(-1, 3).Value = y ' copy y to line above new column

            z = z + 1 ' person id plus 1

            y = 0 ' reset y

    End If:

   

    If ActiveCell.Offset(0, -7).Value = "" Then Exit Do

   

        

    If ActiveCell.Offset(0, -7).Value <> z Then ' if the person id is not plus 1, add another to person id(z)

   

            z = z + 1

    End If:

           ' whole thing will start again down the rows

          

    If ActiveCell.Offset(0, -7).Value = "" Then Exit Do

   

    

Loop Until ActiveCell.Offset(0, -7).Value = "" ' doesn't get to this point at the end of sheet

' need to look at the process and see what we can do

 

'''''''''' TOTALS FOR 12 Week

 

Worksheets("12Wk RECIST").Activate

Columns("O:O").Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .ThemeColor = xlThemeColorAccent4

        .TintAndShade = 0.799981688894314

        .PatternTintAndShade = 0

    End With

    With Selection.Font

        .ColorIndex = xlAutomatic

        .TintAndShade = 0

    End With

   

Range("H2").Activate

z = 1

y = 0

 

 

 

ActiveCell.Offset(-1, 7).Value = "Total Lesion Measurements 12 at Weeks"

 

Do ' start loop

   

    If ActiveCell.Offset(0, -7).Value = "" Then Exit Do

   

    If ActiveCell.Offset(0, -7).Value = z Then ' if pt id = 1 do next

       

        Do

            x = ActiveCell.Value

            y = y + x ' y updated

                       

            ActiveCell.Offset(1, 0).Activate

           

        If ActiveCell.Offset(0, -7).Value = "" Then Exit Do

       

        Loop Until ActiveCell.Offset(0, -7).Value > z ' once person id changes restart

            ActiveCell.Offset(-1, 7).Value = y ' copy y to line above new column

            z = z + 1 ' person id plus 1

            y = 0 ' reset y

    End If:

   

    If ActiveCell.Offset(0, -7).Value = "" Then Exit Do

   

        

    If ActiveCell.Offset(0, -7).Value <> z Then ' if the person id is not plus 1, add another to person id(z)

   

            z = z + 1

    End If:

           ' whole thing will start again down the rows

          

    If ActiveCell.Offset(0, -7).Value = "" Then Exit Do

   

    

Loop Until ActiveCell.Offset(0, -7).Value = "" ' doesn't get to this point at the end of sheet

' need to look at the process and see what we can do

 

 

End Sub

 

 

Sub checkResponse()

'' STARTED THIS GOT TO STEP THROUGH AND WORK OUT WHERE TO GO NEXT

 

Worksheets("12Wk RECIST").Activate

Range("P1").Activate

ActiveCell.Value = "Exp Response From Total Measurements"

 

''''''''''''''''''''''chnge column colour yellow and font black

Columns("P:P").Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .ThemeColor = xlThemeColorAccent4

        .TintAndShade = 0.799981688894314

        .PatternTintAndShade = 0

    End With

    With Selection.Font

        .ColorIndex = xlAutomatic

        .TintAndShade = 0

    End With

'''''''''''''''''''''''''''''''''''''''''

 

 

Range("O2").Activate

 

Worksheets("BL RECIST").Activate

Range("K2").Activate

Dim response As Variant

Dim blTotal As Double

Dim perID As Integer

Dim responseWorked As String

 

perID = 1

 

 

Do '#1

    If ActiveCell.Offset(1, -10).Value = "" Then '*8

    Exit Sub '*8

    End If

           

                Do '#2 LOOP TO FIND NEXT CELL WITH A TOTAL

                    If ActiveCell.Value = "" Then '*1

                        ActiveCell.Offset(1, 0).Activate

                    End If '*1

                Loop Until ActiveCell.Value <> "" '#2

  

   

    Do '#3 FIND OUT IF THE PERSON ID == WHAT WE EXPECT

   

        If ActiveCell.Offset(0, -10).Value = perID Then '*2

       

                      Do '#4 LOOP TO FIND NEXT CELL WITH A TOTAL

                        If ActiveCell.Value = "" Then '*3

                            ActiveCell.Offset(1, 0).Activate

                        End If '*3

                     Loop Until ActiveCell.Value <> "" '#4

           

                    

            blTotal = ActiveCell.Value

            Worksheets("12Wk RECIST").Activate

       

                        Do '#5

                                If ActiveCell.Value = "" Then '*4

                                    ActiveCell.Offset(1, 0).Activate

                                End If '*4

                        Loop Until ActiveCell.Value <> "" '#5 MOVE DOWN UNTIL WE FIND A TOTAL

                Do '#6

               

                    If ActiveCell.Offset(0, -14).Value = perID Then ' '*5 CHECK TO SEE IF THE ID IS THE SAME

                       

                                

                        Do '#7 LOOP TO FIND NEXT CELL WITH A TOTAL

                        If ActiveCell.Value = "" Then '*6

                            ActiveCell.Offset(1, 0).Activate

                        End If '*6

                        Loop Until ActiveCell.Value <> "" '#7

                       

                            response = ActiveCell.Value

           

                            Select Case response ' RAISE A CASE TO CHECK WHICH RESPONSE

           

                            Case Is = 0

                                responseWorked = "Complete Response"

                               

                            Case Is <= (blTotal - (blTotal * 0.3))

                                responseWorked = "Partial Response"

                            

                            Case Is < (blTotal + (blTotal * 0.2))

                                responseWorked = "Stable Disease"

                                                  

                            Case Else

                                responseWorked = "Progressive Disease"

                               

                            End Select

                           

                            ActiveCell.Offset(0, 1) = responseWorked

                            perID = perID + 1

                            Worksheets("BL RECIST").Activate

                            Exit Do

                                          

                    

                    ElseIf ActiveCell.Offset(1, -14).Value = perID Then '*5 I DON'T THINK THIS IS REQUIRED

                    ActiveCell.Offset(1, 0).Activate

                   

                    ElseIf ActiveCell.Offset(1, -14).Value > perID Then '*5

                    perID = ActiveCell.Offset(1, -14).Value

                    Worksheets("BL RECIST").Activate

                    Exit Do

                   

                    Else '*5

                        ActiveCell.Offset(1, 0).Activate

                     

                        

                     If ActiveCell.Offset(0, -14).Value <> perID Then '*7

                       ActiveCell.Offset(1, 0).Activate

                    End If '*7

                            End If '*5

           

                Loop Until ActiveCell.Offset(1, -14).Value = "" '#6

               

        ElseIf ActiveCell.Offset(1, -10).Value > perID Then '*2

            perID = perID + 1

        Else '*2

            ActiveCell.Offset(1, 0).Activate

        End If '*2

       

        

    Loop Until ActiveCell.Offset(1, -10).Value = "" '#3

 

    Exit Sub

Loop Until ActiveCell.Offset(1, -10).Value = "" '#1

 

Exit Sub

End Sub

 

Sub checkExpTargetRespVeCRF()

 

Worksheets("12Wk RECIST").Activate

 

Range("P2").Activate

Dim expResponse As String

Dim eCRFResponse As String

Dim result As String

 

'Application.ScreenUpdating = False

 

 

Do '#3

    Do '#2

   

        Do '#1

            If ActiveCell.Value = "" Then '*1

                ActiveCell.Offset(1, 0).Activate

            End If

        Loop Until ActiveCell.Value <> "" '#1

       

        expResponse = ActiveCell.Value

        eCRFResponse = ActiveCell.Offset(0, -6).Value

       

                    Select Case eCRFResponse

                        Case Is = ""

                            result = "Missing"

                       

                        Case Is = expResponse

                            result = "Match"

                        

                        Case Is = "N/A"

                            result = "N/A"

                       

                        Case Is = "Not Evaluable"

                            result = "Not Evaluable"

                           

                        Case Else

                            result = "No Match"

                    End Select

           

                If result = "Missing" Then '*2

                ActiveCell.Offset(1, 0).Activate

                Exit Do '#2

                

                ElseIf result = "N/A" Then '*2

                ActiveCell.Offset(1, 0).Activate

                Exit Do '#2

               

                ElseIf result = "Not Evaluable" Then '*2

                ActiveCell.Offset(1, 0).Activate

                Exit Do '#2

               

                ElseIf result = "Match" Then '*2

                ActiveCell.Offset(1, 0).Activate

                Exit Do '#2

               

                Else '*2

                With Selection.Interior

                    .Pattern = xlSolid

                    .PatternColorIndex = xlAutomatic

                    .Color = 255

                    .TintAndShade = 0

                    .PatternTintAndShade = 0

                End With

                ActiveCell.Offset(1, 0).Activate

                End If '*2

               

    Loop Until ActiveCell.Offset(1, -15).Value = "" '#2

       

Loop Until ActiveCell.Offset(1, -15).Value = "" '#3

End Sub

 

 

Sub NTLatBL2()

 

Dim perID As Integer

Dim NTL As String

 

rowQColourandTitle

 

Range("K2").Activate

 

Worksheets("BL RECIST").Activate

Range("J2").Activate

 

Do '#1

    perIDBL = ActiveCell.Offset(0, -9)

    NTL = ActiveCell.Value

   

    If ActiveCell.Offset(0, -9).Value = "" Then

       

        Exit Sub

    End If

   

    If NoMoreLines = 1 Then

       

        Exit Sub

    End If

   

    If NTL = "No" Then '$1

        ActiveCell.Offset(1, 0).Activate

       

    ElseIf NTL = "Yes" Then '$1

   

            If ActiveCell.Offset(0, 1) <> "" Then

            TwelveWeekMatch 'SUB

            Else

            ActiveCell.Offset(1, 0).Activate

            End If

                   

    ElseIf NTL = "" Then '$1

        ActiveCell.Offset(1, 0).Activate

   

    Else '$1

        ActiveCell.Offset(1, 0).Activate

        MsgBox ("Error")

        Exit Sub

       

    End If '$1

   

Loop '#1

 

End Sub

 

 

Sub TwelveWeekMatch()

 

Worksheets("12Wk RECIST").Activate

Dim perID12Wk As Integer

 

Do

 

    perID12Wk = ActiveCell.Offset(0, -10).Value

   

    Select Case perID12Wk

       

        Case Is = 0

            NoMoreLines = 1

            Exit Sub

           

        Case Is = perIDBL

            Row 'SUB

                If RowYN = 1 Then 

                    ActiveCell.Offset(1, 0).Activate

                    Worksheets("BL RECIST").Activate

                    ActiveCell.Offset(1, 0).Activate

                    RowYN = 0

                Exit Sub

               

                Else

                ActiveCell.Offset(1, 0).Activate

                End If

                                       

        Case Is < perIDBL

            ActiveCell.Offset(1, 0).Activate

           

        Case Is > perIDBL And perIDBL > 0

            Worksheets("BL RECIST").Activate

            ActiveCell.Offset(1, 0).Activate

            Exit Sub

                            

        Case Else

            Exit Sub

           

    End Select

   

Loop

 

End Sub

 

Sub Row()

 

    If ActiveCell.Offset(0, 4) <> "" Then

        NTLat12Wk 'SUB

        RowYN = 1

        Exit Sub

   

    Else

   

    End If

   

End Sub

 

Sub NTLat12Wk()

Dim TwWkNTL As String

Dim TwWkNTLResp As String

 

TwWkNTL = ActiveCell.Offset(0, 1).Value

TwWkNTLResp = ActiveCell.Value

 

    If TwWkNTL = "No" Then

   

        If TwWkNTLResp = "Complete Response" Then

            Exit Sub

        Else

            ActiveCell.Offset(0, 6).Value = "Needs to be Checked - TARGET LESION RESPONSE AT 12 WEEKS IS INCORRECT - NTL'S ARE MISSING FROM THE eCRF"

            ColourCellRed 'SUB

            Exit Sub

        End If

       

    ElseIf TwWkNTL = "Yes" Then

       

        ActiveCell.Offset(0, 6).Value = "Needs to be Checked - TARGET LESION RESPONSE AT 12 WEEKS NEEDS MANUAL CHECKS"

        ColourCellRed 'SUB

        Exit Sub

    Else

        ActiveCell.Offset(0, 6).Value = "Missing Data"

        Exit Sub

 

    End If

   

End Sub

 

 

Sub rowQColourandTitle()

Worksheets("12Wk RECIST").Activate

Columns("Q:Q").ColumnWidth = 36.71

Columns("Q:Q").Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .ThemeColor = xlThemeColorAccent4

        .TintAndShade = 0.799981688894314

        .PatternTintAndShade = 0

    End With

    With Selection.Font

        .ColorIndex = xlAutomatic

        .TintAndShade = 0

    End With

Range("Q1").Activate

ActiveCell.Value = "12 Week NTL"

End Sub

 

Sub ColourCellRed()

 

    ActiveCell.Offset(0, 6).Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .Color = 255

        .TintAndShade = 0

        .PatternTintAndShade = 0

    End With

    ActiveCell.Offset(0, -6).Select

End Sub

 

Sub newlesions()

 

rowQColourandTitle2

 

Range("L2").Activate

 

Do

    If ActiveCell.Offset(0, -11).Value = "" Then

        Exit Sub

    End If

   

    If ActiveCell.Offset(0, 3) <> "" Then

        checkforlesions1

        ActiveCell.Offset(1, 0).Activate

    Else

        ActiveCell.Offset(1, 0).Activate

    End If

   

Loop

 

End Sub

 

 

Sub checkforlesions1()

 

    If ActiveCell.Value = "Yes" Then

        ActiveCell.Offset(0, 6).Value = "Yes"

       

    Else

    End If

   

End Sub

 

 

Sub rowQColourandTitle2()

Worksheets("12Wk RECIST").Activate

Columns("R:R").ColumnWidth = 15

Columns("R:R").Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .ThemeColor = xlThemeColorAccent4

        .TintAndShade = 0.799981688894314

        .PatternTintAndShade = 0

    End With

    With Selection.Font

        .ColorIndex = xlAutomatic

        .TintAndShade = 0

    End With

Range("R1").Activate

ActiveCell.Value = "New Lesions at 12 Wks?"

End Sub

 

Sub NewLesionListed()

 

rowQColourandTitle3

Worksheets("12Wk RECIST").Activate

Range("R2").Activate

 

Do

 

    If ActiveCell.Offset(0, -11).Value = "" Then

        Exit Sub

    End If

 

   

    If ActiveCell.Value <> "" Then

    checkforlesions2

    ActiveCell.Offset(1, 0).Activate

   

    

    Else

    ActiveCell.Offset(1, 0).Activate

 

    End If

 

Loop

 

End Sub

 

Sub checkforlesions2()

 

    If ActiveCell.Offset(0, -5) <> "" Then

    ActiveCell.Offset(0, 1).Value = "Yes"

    Exit Sub

   

    Else

        ColourCellRed2

        ActiveCell.Offset(0, 1).Value = "New Lesions have not been listed"

     End If

    

End Sub

Sub rowQColourandTitle3()

Worksheets("12Wk RECIST").Activate

Columns("S:S").ColumnWidth = 15

Columns("S:S").Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .ThemeColor = xlThemeColorAccent4

        .TintAndShade = 0.799981688894314

        .PatternTintAndShade = 0

    End With

    With Selection.Font

        .ColorIndex = xlAutomatic

        .TintAndShade = 0

    End With

Range("S1").Activate

ActiveCell.Value = "New Lesions Listed?"

End Sub

 

Sub ColourCellRed2()

 

    ActiveCell.Offset(0, 1).Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .Color = 255

        .TintAndShade = 0

        .PatternTintAndShade = 0

    End With

    ActiveCell.Offset(0, -1).Select

End Sub

 

Sub expectedoverallresponse()

 

rowQColourandTitle4

Worksheets("12Wk RECIST").Activate

Range("N2").Activate

 

Do

    If ActiveCell.Offset(0, -13).Value = "" Then

        Exit Sub

    End If

 

findline

ActiveCell.Offset(1, 0).Activate

 

Loop

 

End Sub

 

Sub findline()

    If ActiveCell.Offset(0, 1).Value = "" Then

    Exit Sub

   

    Else

    expreponse

    Exit Sub

    End If

   

        

End Sub

 

Sub expreponse()

Dim overallresponse As String

Dim tumourdiameter As Long

 

 

overallresponse = ActiveCell.Value

tumourdiameter = ActiveCell.Offset(0, 1).Value

 

    Select Case tumourdiameter

    Case Is = 0

        If overallresponse = "Complete Response" Then

        expoverallresponse = "Complete Response"

        checking1

       

        

        ElseIf overallreponse = "" Then

        expoverallresponse = "Missing Data"

        checking1

        Exit Sub

               

        Else

        ColourCellRed5

        expoverallresponse1st = "Check Overall Response"

        Exit Sub

        End If

       

    Case Is > 0

        expoverallresponse = ActiveCell.Offset(0, 2).Value

        checking1

    End Select

    Exit Sub

   

 

End Sub

 

Sub checking1()

Dim ntlresponse As String

 

    If ActiveCell.Offset(0, 3).Value = "" Then

    checking2

    Exit Sub

        

    Else

   

        ntlresponse = ActiveCell.Offset(0, -3).Value

           

        Select Case ntlresponse

            

             Case Is = "Not Evaluable"

              expoverallresponse = "Not Evaluable"

             

             Case Is = "N/A"

             expoverallresponse = "Check Overall Response, non taget lesion response is not consistent with the overall response"

             ColourCellRed5

            

             Case Is = "No Complete Response/ No Progressive Disease"

                If ActiveCell.Offset(0, 2) <> "Partial Response" Then

                    expoverallresponse = "Stable Disease"

                Else

                    expoverallresponse = "Partial Response"

                End If

               

             Case Is = "Complete Response"

             expoverallresponse = "Complete Response"

              

             Case Is = "Progressive Disease"

             expoverallresponse = "Progressive Disease"

            

             Case Is = ""

             expoverallresponse = "Missing Data"

                        

             Case Else

             expoverallresponse = "Check Overall Response"

             ColourCellRed5

       

        End Select

   

    checking2

        Exit Sub

    End If

   

End Sub

 

Sub checking2()

 

Dim newlesions As String

 

    If ActiveCell.Offset(0, 4).Value = "" Then

    ActiveCell.Offset(0, 6).Value = expoverallresponse1st & expoverallresponse

    Exit Sub

       

    Else

        newlesions = ActiveCell.Offset(0, 4)

            

            Select Case newlesions

       

            Case Is = "Yes"

            expoverallresponse = "Progressive Disease"

            ActiveCell.Offset(0, 6).Value = expoverallresponse1st & expoverallresponse

           

            Case Else

            expoverallresponse = "Check Responses for this pt"

            ActiveCell.Offset(0, 6).Value = expoverallresponse1st & expoverallresponse

           

            End Select

        Exit Sub

       

    End If

   

End Sub

Sub rowQColourandTitle4()

Worksheets("12Wk RECIST").Activate

Columns("T:T").ColumnWidth = 15

Columns("T:T").Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .ThemeColor = xlThemeColorAccent4

        .TintAndShade = 0.799981688894314

        .PatternTintAndShade = 0

    End With

    With Selection.Font

        .ColorIndex = xlAutomatic

        .TintAndShade = 0

    End With

Range("T1").Activate

ActiveCell.Value = "Expected Overall Response"

End Sub

 

Sub ColourCellRed5()

 

    ActiveCell.Offset(0, 6).Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .Color = 255

        .TintAndShade = 0

        .PatternTintAndShade = 0

    End With

    ActiveCell.Offset(0, -6).Select

End Sub

 

Sub CheckExpectedRespAgainsteCRF()

 

 

 

Worksheets("12Wk RECIST").Activate

Range("T2").Activate

Do

 

    If ActiveCell.Offset(0, -15).Value = "" Then

        Exit Sub

    End If

   

        If ActiveCell.Offset(0, -5).Value = "" Then

            ActiveCell.Offset(1, 0).Activate

       

        Else

            checkresps

            ActiveCell.Offset(1, 0).Activate

 

         End If

 

Loop

 

 

End Sub

 

Sub checkresps()

 

Dim expresp As String

Dim CRFresp As String

 

        expresp = ActiveCell.Value

        CRFresp = ActiveCell.Offset(0, -6).Value

       

        If expresp = CRFresp Then

            ActiveCell.Offset(1, 0).Activate

       

        ElseIf expresp = "Missing Data" Then

            ActiveCell.Offset(1, 0).Activate

       

        Else

            ColourCellRed6

            ActiveCell.Value = ActiveCell.Value & " - Check Response"

            ActiveCell.Offset(1, 0).Activate

        End If

 

End Sub

Sub ColourCellRed6()

 

    ActiveCell.Select

    With Selection.Interior

        .Pattern = xlSolid

        .PatternColorIndex = xlAutomatic

        .Color = 255

        .TintAndShade = 0

        .PatternTintAndShade = 0

    End With

End Sub

 

 

Sub wraptext()

 

    Worksheets("BL RECIST").Activate

    Columns("K:K").Select

    With Selection

        .HorizontalAlignment = xlGeneral

        .VerticalAlignment = xlBottom

        .wraptext = True

        .Orientation = 0

        .AddIndent = False

        .IndentLevel = 0

        .ShrinkToFit = False

       .ReadingOrder = xlContext

        .MergeCells = False

    End With

 

 

    Worksheets("12Wk RECIST").Activate

    Columns("O:T").Select

    With Selection

        .HorizontalAlignment = xlGeneral

        .VerticalAlignment = xlBottom

        .wraptext = True

        .Orientation = 0

        .AddIndent = False

        .IndentLevel = 0

        .ShrinkToFit = False

        .ReadingOrder = xlContext

        .MergeCells = False

    End With

End Sub

 

 

Sub deletesheets()

 

Application.DisplayAlerts = False

 

    Sheets("Cervical CA Status").Select

    ActiveWindow.SelectedSheets.Delete

   

    Sheets("12 Week Dates").Select

    ActiveWindow.SelectedSheets.Delete

   

Application.DisplayAlerts = True

 

End Sub

 

Sub deleteconnections()

Dim cn As WorkbookConnection

Dim qr As WorkbookQuery

On Error Resume Next

For Each cn In ThisWorkbook.Connections

    cn.Delete

Next

For Each qr In ThisWorkbook.Queries

    qr.Delete

Next

End Sub

 

 

 

'''''''''''''''DELETE BUTTON'''''

Sub deletebutton()

 

    Worksheets("BL RECIST").Activate

    ActiveSheet.Shapes.Range(Array("Rectangle: Rounded Corners 1")).Select

    Selection.Delete

    Worksheets("12Wk RECIST").Activate

 

End Sub