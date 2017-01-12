VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Risk Assessment Form"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13500
   OleObjectBlob   =   "UserForm1_Intensity.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnPopMskIntensities_Click()
'Dim TheSheet As Worksheet
'Set TheSheet = Sheets("LookUpList")
MsgBox "count of msk" & Application.CountA(Range("ES:ES"))
For i = 1 To Application.CountA(Range("ES:ES"))
If Len(Range("ES" & i)) > 0 Then
comboMsk.AddItem (Range("ES" & i))
End If
Next

End Sub

Private Sub checkValue_Click()
'''MsgBox "hi"
''MsgBox txtBox.Text
Dim WB As Workbook
Dim SourceWB As Workbook
Dim WS As Worksheet
Dim ASheet As Worksheet

'Turns off screenupdating and events:
Application.ScreenUpdating = False
Application.EnableEvents = False

'Sets the variables:
Set WB = ActiveWorkbook
Set ASheet = ActiveSheet

''''''''''''''
'filter = "Text files (*.dbf),*.dbf"
'caption = "Please Select an input file "
'customerFilename = Application.GetOpenFilename(filter, caption)

'Set customerWorkbook = Application.Workbooks.Open(customerFilename)
''''''''''''''

Set SourceWB = Application.Workbooks.Open(Application.GetOpenFilename("All Files (*.*),*.*"))   'Modify to match

'Copies each sheet of the SourceWB to the end of original wb:
For Each WS In SourceWB.Worksheets
 If Application.CountA(Cells) <> 0 Then
        'MsgBox ActiveSheet.Name & " is not empty"
        ''MsgBox "hey"
        WS.Copy after:=WB.Sheets(WB.Sheets.count)
        ''MsgBox "hi"
      
   
    
Dim iRows As Long
Dim iCols As Long
Dim fa As Double
Dim fpa As Double
Dim faFpaRatio As Double
Dim polygonId As Long
Dim polygonIdPlusOne As Long
Dim arrayStore(3000) As Variant
Dim arraySum(3000) As Variant
Dim arrayForRatio()
Dim valueArr As Long
Dim Avg As Double
Dim countNonZero As Long
Dim countFinal As Long
Dim arrCollect(3000) As Variant
Dim colNameFA As String
Dim colNameFPA As String
Dim colNameTFID As String
Dim colNameMBT As String
Dim colNameOCCD As String
Dim colNameOCCN As String
Dim clusterNo As Integer
Dim loopStrtPt  As Integer
Dim y As Integer
Dim sumOfIndvFA As Double
Dim arrayMBT(3000) As Variant
Dim arrayMBTposition(3000) As Variant
Dim arrayFA(3000) As Variant
Dim returnValGetUnique(1200, 300) As Variant
Dim arrRowNoMatch(3000) As Variant
Dim IndvFAForEachMBT As Double
Dim contributionFactore As Double
Dim sumOfFAforContrbF As Double
Dim contrbFAperMBT As Double
Dim indvOccDayForEachMBT As Double
Dim indvOccNightForEachMBT As Double
Dim arrayOccDayMBT(3000) As Variant
Dim arrayOccNightMBT(3000) As Variant
Dim contrbOccDayMBT As Double
Dim contrbOccNightMBT As Double
Dim countClusterObjects As Integer
Dim sumFaFpaRatio As Integer



clusterNo = 0
valueArr = 0
Avg = 0
countNonZero = 0
arraySum(0) = 0
''MsgBox "active wb name" & WB.Sheets.Count

''MsgBox "F1" & Cells(1, 6)


iRows = Application.CountA(Range("A:A"))
''MsgBox iRows

iCols = Application.CountA(Range("1:1"))
''MsgBox iCols

'Check Column name of TargetFID (polygonID) Start
 

    Dim strSearchTFID As String
    Dim aCellTFID As Range
    strSearchTFID = "TARGET_FID"

    Set aCellTFID = ActiveSheet.Rows(1).Find(What:=strSearchTFID, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    If Not aCellTFID Is Nothing Then
      '  'MsgBox "Value Found in Cell " & aCellTFID.Address & _
       ' " and the Cell Column Number is " & aCellTFID.Column
    End If
     ColNo = aCellTFID.Column
     colNameTFID = Split(Cells(, ColNo).Address, "$")(1)
   '  'MsgBox "column name of TFID"
   '  'MsgBox colNameTFID
     
'Check Column name of TargetFID (polygonID) Start End



'Check Column name of FA Start
 

    Dim strSearchFA As String
    Dim aCellFA As Range

    strSearchFA = "TFA"

    Set aCellFA = ActiveSheet.Rows(1).Find(What:=strSearchFA, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    If Not aCellFA Is Nothing Then
     '    'MsgBox "Value Found in Cell " & aCellFA.Address & _
     '   " and the Cell Column Number is " & aCellFA.Column
    End If
     ColNo = aCellFA.Column
     colNameFA = Split(Cells(, ColNo).Address, "$")(1)
     ''MsgBox "column name of FA"
     ''MsgBox colNameFA
     
'Check column name of FA End

'Check Column name of FPA Start


    Dim strSearchFPA As String
    Dim aCellFPA As Range

    strSearchFPA = "FPA"

    Set aCellFPA = ActiveSheet.Rows(1).Find(What:=strSearchFPA, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    If Not aCellFPA Is Nothing Then
      '  'MsgBox "Value Found in Cell " & aCellFPA.Address & _
        " and the Cell Column Number is " & aCellFPA.Column
    End If
     ColNo = aCellFPA.Column
     colNameFPA = Split(Cells(, ColNo).Address, "$")(1)
     ''MsgBox "column name of FPA"
     ''MsgBox colNameFPA
     
'Check column name of FPA End

'Check Column name of MBT Start


    Dim strSearchMBT As String
    Dim aCellMBT As Range

    strSearchMBT = "MBT"

    Set aCellMBT = ActiveSheet.Rows(1).Find(What:=strSearchMBT, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    If Not aCellMBT Is Nothing Then
       'MsgBox "Value Found in Cell " & aCellFPA.Address & _
        " and the Cell Column Number is " & aCellMBT.Column
    End If
     ColNo = aCellMBT.Column
     colNameMBT = Split(Cells(, ColNo).Address, "$")(1)
     ''MsgBox "column name of FPA"
     ''MsgBox colNameFPA
     
'Check column name of MBT End

''Check Column name of Occupancy Day Start


    Dim strSearchOCCD As String
    Dim aCellOCCD As Range

    strSearchOCCD = "occD"

    Set aCellOCCD = ActiveSheet.Rows(1).Find(What:=strSearchOCCD, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    If Not aCellOCCD Is Nothing Then
      '  'MsgBox "Value Found in Cell " & aCellFPA.Address & _
        " and the Cell Column Number is " & aCellFPA.Column
    End If
     ColNo = aCellOCCD.Column
     colNameOCCD = Split(Cells(, ColNo).Address, "$")(1)
     ''MsgBox "column name of FPA"
     ''MsgBox colNameFPA

''Check column name of Occupancy Day End

'Check Column name of Occupancy Night Start


    Dim strSearchOCCN As String
    Dim aCellOCCN As Range

    strSearchOCCN = "occN"

    Set aCellOCCN = ActiveSheet.Rows(1).Find(What:=strSearchOCCN, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    If Not aCellOCCN Is Nothing Then
      '  'MsgBox "Value Found in Cell " & aCellFPA.Address & _
        " and the Cell Column Number is " & aCellFPA.Column
    End If
     ColNo = aCellOCCN.Column
     colNameOCCN = Split(Cells(, ColNo).Address, "$")(1)
     ''MsgBox "column name of FPA"
     ''MsgBox colNameFPA
     
'Check column name of Occupancy Day End

Dim mbtCount As Integer
''''''''''''''''''''
For x = 2 To iRows

Range("AZ" & x) = (Range(colNameFA & x)) / (Range(colNameFPA & x))


''' '''' ''''
If (Range(colNameTFID & x + 1) <> Range(colNameTFID & x)) Then
 If (Range(colNameTFID & x) <> Range(colNameTFID & x - 1)) Then
''' '''' ''''
countClusterObjects = 1
y = y + 1
''MsgBox "row no : " & y
'MsgBox "single valued cluster"
'MsgBox "targetFID value is : " & Range(colNameTFID & x).Value
arrayMBT(0) = Range(colNameMBT & x).Value
For l = 196 To 264
 If Range(colNameMBT & x).Value = Application.ActiveWorkbook.Worksheets(1).Range("B" & l) Then
 'MsgBox "matched at " & l
 'MsgBox "FA before : " & Application.ActiveWorkbook.Worksheets(1).Cells(l, 9)
 Application.ActiveWorkbook.Worksheets(1).Cells(l, 9) = Range(colNameFA & x).Value + Application.ActiveWorkbook.Worksheets(1).Cells(l, 9)
 'MsgBox "FA after : " & Application.ActiveWorkbook.Worksheets(1).Cells(l, 9)
     
 'MsgBox "occD before : " & Application.ActiveWorkbook.Worksheets(1).Cells(l, 12)
 Application.ActiveWorkbook.Worksheets(1).Cells(l, 12) = Range(colNameOCCD & x).Value + Application.ActiveWorkbook.Worksheets(1).Cells(l, 12)
 'MsgBox "occD after : " & Application.ActiveWorkbook.Worksheets(1).Cells(l, 12)
 
 Application.ActiveWorkbook.Worksheets(1).Cells(l, 13) = Range(colNameOCCN & x).Value + Application.ActiveWorkbook.Worksheets(1).Cells(l, 13)
 End If
Next
    

End If
End If

''''''''''''''''''''




If (Range(colNameTFID & x + 1) = Range(colNameTFID & x)) Then

'MsgBox "targetFID value is : " & Range(colNameTFID & x).Value
'MsgBox "FA/FPA ratio " & Range("AZ" & x)

If Range(colNameMBT & x).Value <> 0 Then
'MsgBox "not zero at :: " & x
''MsgBox "mbt count is :: " & mbtCount
arrayMBT(mbtCount) = Range(colNameMBT & x).Value
arrayFA(mbtCount) = Range(colNameFA & x)
arrayOccDayMBT(mbtCount) = Range(colNameOCCD & x)
arrayOccNightMBT(mbtCount) = Range(colNameOCCN & x)
arrRowNoMatch(mbtCount) = x
mbtCount = mbtCount + 1
countNonZero = countNonZero + 1
'sumFaFpaRatio = Range("AZ" & x) + sumFaFpaRatio
arrayStore(x) = Range("AZ" & x)

End If

'arrayStore(x) = Range("AZ" & x)


arrCollect(x) = x

''MsgBox "arrCollect"
''MsgBox arrCollect(x)

If arrayStore(x) <> 0 Then
'countNonZero = countNonZero + 1

''MsgBox "arraystore(x) in first loop"
''MsgBox arrayStore(x)
'MsgBox "countNonzero is : " & countNonZero

End If

End If

If (Range(colNameTFID & x + 1) <> Range(colNameTFID & x)) Then
 If (Range(colNameTFID & x) = Range(colNameTFID & x - 1)) Then
 
'MsgBox "targetFID value is : " & Range(colNameTFID & x).Value
'MsgBox "FA/FPA ratio " & Range("AZ" & x)

If Range(colNameMBT & x).Value <> 0 Then
'MsgBox "not zero at :: " & x
''MsgBox "mbt count is :: " & mbtCount
arrayMBT(mbtCount) = Range(colNameMBT & x).Value
arrayFA(mbtCount) = Range(colNameFA & x)
arrayOccDayMBT(mbtCount) = Range(colNameOCCD & x)
arrayOccNightMBT(mbtCount) = Range(colNameOCCN & x)
arrRowNoMatch(mbtCount) = x

mbtCount = mbtCount + 1
arrayStore(x) = Range("AZ" & x)

'sumFaFpaRatio = Range("AZ" & x) + sumFaFpaRatio
End If
'MsgBox "countFinal is : " & countFinal
'arrayStore(x) = Range("AZ" & x)

''MsgBox "arrastore(x) in second loop"
''MsgBox arrayStore(x)

If arrayStore(x) <> 0 Then
''MsgBox "countNonzero"
''MsgBox countNonZero

countFinal = countNonZero + 1
'MsgBox "countFinal is : " & countFinal

End If

If arrayStore(x) = 0 Then
''MsgBox "countNonzero"
''MsgBox countNonZero

countFinal = countNonZero
''MsgBox "countFinal " & countFinal

If countFinal = 0 Then
''MsgBox "There is no representative building for the polygon ID"
''MsgBox Range(colNameTFID & x)
End If
End If
''MsgBox "countFinal is : " & countFinal
clusterNo = clusterNo + 1
'MsgBox "cluster No " & clusterNo

'MsgBox "arrayStore count : " & Application.CountA(arrayStore)
'MsgBox "arraySum : " & Application.WorksheetFunction.Sum(arrayStore) & " countFinal " & countFinal
If countFinal > 0 Then
Avg = Application.WorksheetFunction.Sum(arrayStore) / countFinal
End If
'MsgBox "avg is : " & Avg

arrCollect(x) = x

''MsgBox "arrCollect in second loop"
''MsgBox arrCollect(x)
Erase arrayStore

''''''''''''''''''
''MsgBox "size of arrCollect" & Application.CountA(arrCollect)
''''''''''''''''''

End If

''''''''

If clusterNo = 1 Then
loopStrtPt = 1
Else
'MsgBox "loopStrtPt : " & y + 1
loopStrtPt = y + 1
End If

''MsgBox "arrMBT count is : " & Application.CountA(arrayMBT)
'MsgBox "NON UNIQUE MBT count is : " & Application.CountA(arrayMBT)
''MsgBox "occDay count is : " & Application.CountA(arrayOccDayMBT)
For k = 0 To (Application.CountA(arrayOccDayMBT) - 1)
'MsgBox "arrayOccDayMBT values++++++++ " & arrayOccDayMBT(k)
Next

''' for summation of indiv floor areas against each cluster
sumOfFAforContrbF = 0
''MsgBox "average is :: " & Avg

For m = 1 To Application.CountA(arrCollect)
y = loopStrtPt
''MsgBox " loop y is :" & y + 1
'MsgBox "value of FA : " & Range(colNameFA & y + 1).Value & " SumofFA is : " & sumOfFAforContrbF

sumOfFAforContrbF = sumOfFAforContrbF + Range(colNameFA & y + 1)

Range(colNameFA & y + 1) = Avg * Range(colNameFPA & y + 1)
''MsgBox "FPA " & Range(colNameFPA & y + 1) & " avg " & Avg & " each indv FA is : " & Range(colNameFA & y + 1)

sumOfIndvFA = sumOfIndvFA + Range(colNameFA & y + 1)
loopStrtPt = loopStrtPt + 1
Next
''MsgBox "sum total of IndvFA for the cluster is ::" & sumOfIndvFA
'MsgBox "sumOfFAforContrbF +++::" & sumOfFAforContrbF
''''''''
''''''''

'' Code for matching the MBT values
Dim arr As New Collection, a
  Dim i As Long
  Dim count As Integer
  
  On Error Resume Next
  
  For Each a In arrayMBT
     arr.Add a, a
  Next
'MsgBox "UNIQUE MBT Count : " & arr.count

  For i = 1 To arr.count
'  'MsgBox "unique values are :: " & arr(i)
  Next
 ' 'MsgBox "arr count : " & arr.count
  For Z = 1 To Application.CountA(arrayMBT)
  ''MsgBox "arrMBT value s " & arrayMBT(Z)
  Next
 'count = arr.count
  'getUnique = arr
Dim countMBT As Integer
' For n = 0 To Application.CountA(arrayMBT)
' If arrayMBT(n) = "AM11" Then
' countMBT = countMBT + 1
' End If
'  Next
 
 Dim mbt As String
 Dim arrMbt
 Dim faMBT As Double
 Dim occDay As Double
 Dim occNight As Double
 
 
' 'MsgBox "count of unique values :: " & arr.count

 For b = 1 To ((arr.count + 1))
 ''MsgBox "b is : " & b - 1
 countMBT = 0
 IndvFAForEachMBT = 0
 mbt = arr(b - 1)

  For c = 0 To ((Application.CountA(arrayMBT) - 1))
  arrMbt = arrayMBT(c)
   
   ' 'MsgBox "b = " & (b - 1) & " arr : " & arr(b - 1) & " c = " & (c) & " arrMBT : " & arrayMBT(c)
  ''MsgBox "countMBT outta IF "
   If arrMbt = mbt Then
   'MsgBox "match from : " & arr(b - 1) & " match to: " & arrayMBT(c) & " match at : " & (c) & " corresponding FA is " & arrayFA(c) & " occDay is: " & arrayOccDayMBT(c) & " occNight is : " & arrayOccNightMBT(c)
   ''MsgBox "FA is " & arrayFA(n)
   faMBT = arrayFA(c)
   occDay = arrayOccDayMBT(c)
   occNight = arrayOccNightMBT(c)
   IndvFAForEachMBT = IndvFAForEachMBT + faMBT
   countMBT = countMBT + 1
   indvOccDayForEachMBT = indvOccDayForEachMBT + occDay
   indvOccNightForEachMBT = indvOccNightForEachMBT + occNight
   End If
   Next
   ' 'MsgBox " OUTSIDE <> 0 b4 is : " & b - 1
    contributionFactore = (IndvFAForEachMBT) / (sumOfFAforContrbF)
    'MsgBox "contrb Factor for  " & arr(b - 1) & " is : " & contributionFactore
   ' If contributionFactore <> 0 Then
   If (b - 1) > 0 Then
     
    ''MsgBox " inside <> 0 b4 is : " & b - 1
    ''MsgBox "final counter for  " & arr(b - 1) & " is : " & countMBT
    
    ''MsgBox " sum of Indv FA for each mbt " & IndvFAForEachMBT
    'MsgBox "occD Total :" & indvOccDayForEachMBT & " occ Night Total : " & indvOccNightForEachMBT
    'MsgBox "occ Ratio Day : " & indvOccDayForEachMBT / IndvFAForEachMBT & " occ Ratio Night : " & indvOccNightForEachMBT / IndvFAForEachMBT
    'MsgBox "sumOfFAforContrbF : " & sumOfFAforContrbF
    
    ''MsgBox "sum total of all indv FAs is " & sumOfIndvFA
    
    contrbFAperMBT = sumOfIndvFA * contributionFactore
    'MsgBox "contribution of " & arr(b - 1) & " is : " & contrbFAperMBT
    
    contrbOccDayMBT = (indvOccDayForEachMBT / IndvFAForEachMBT) * contrbFAperMBT
    contrbOccNightMBT = (indvOccNightForEachMBT / IndvFAForEachMBT) * contrbFAperMBT
    'MsgBox "occ D final: " & contrbOccDayMBT & " contrb OccN final : " & contrbOccNightMBT
    '' for putting values into indv cells
     For l = 126 To 161
     
     ''MsgBox "range BBBBBBBB+++++++ " & Application.ActiveWorkbook.Worksheets(1).Range("B" & l)
    If arr(b - 1) = Application.ActiveWorkbook.Worksheets(1).Range("B" & l) Then
    '
    'MsgBox "gotta " & arr(b - 1) & " at row no: " & l
    
    ''FA value susbtituting start
     Application.ActiveWorkbook.Worksheets(1).Cells(l, 9) = contrbFAperMBT + Application.ActiveWorkbook.Worksheets(1).Cells(l, 9)
    ''FA value susbtituting start
    ''Occ Day value substituting start
    
    ''MsgBox "before cell value " & Application.ActiveWorkbook.Worksheets(1).Cells(l, 12)
    Application.ActiveWorkbook.Worksheets(1).Cells(l, 12) = contrbOccDayMBT + Application.ActiveWorkbook.Worksheets(1).Cells(l, 12)
    ' 'MsgBox "after cell value " & Application.ActiveWorkbook.Worksheets(1).Cells(l, 12)
    
    ''Occ Day value substituting end
    
    ''Occ Night value substituting start
    
   '' 'MsgBox "before cell value occN " & Application.ActiveWorkbook.Worksheets(1).Cells(l, 13)
    Application.ActiveWorkbook.Worksheets(1).Cells(l, 13) = contrbOccNightMBT + Application.ActiveWorkbook.Worksheets(1).Cells(l, 13)
     ''MsgBox "after cell value occN" & Application.ActiveWorkbook.Worksheets(1).Cells(l, 13)
     
    ''Occ Night value substituting end
    End If
    Next
    '' for putting values into indv cells
    End If
    
    indvOccDayForEachMBT = 0
    indvOccNightForEachMBT = 0
     occDay = 0
     occNight = 0
 Next
  '' Code for matching the MBT values

For p = 0 To ((Application.CountA(arrRowNoMatch)) - 1)
'MsgBox "row no at match :: " & arrRowNoMatch(p)
Next

countClusterObjects = 0
loopStrtPt = 0
contributionFactore = 0
faMBT = 0
occDay = 0
occNight = 0
mbtCount = 0
countFinal = 0
countNonZero = 0
sumOfIndvFA = 0
contributionFactore = 0
IndvFAForEachMBT = 0
sumOfFAforContrbF = 0
Erase arrayOccDayMBT
Erase arrayOccNightMBT
Erase arrCollect
Erase arrayMBT
Erase arrayFA
Erase arrRowNoMatch
Set arr = Nothing

End If


'mbtCount = mbtCount + 1
Next
''MsgBox "total sum"
''MsgBox Application.WorksheetFunction.Sum(Range("B:B"))
Range("F1") = Application.WorksheetFunction.Sum(Range("B:B"))
'Application.ActiveWorkbook.Worksheets(1).Cells(34, 9) = Application.WorksheetFunction.Sum(Range("B:B"))
''''''''''''INPUT
'Application.ActiveWorkbook.Worksheets(1).Cells(8, 11) = txtBox.Text
'Application.ActiveWorkbook.Worksheets(1).Cells(10, 11) = txtSs.Text  ' putting value of Ss
' Application.ActiveWorkbook.Worksheets(1).Cells(11, 11) = txtS1.Text ' Putting value of S1
' Application.ActiveWorkbook.Worksheets(1).Cells(12, 11) = txtTl.Text  ' Putting value of Long Period TL
 
Application.ActiveWorkbook.Worksheets(1).Cells(8, 13) = comboMsk.SelText ' Putting value of MSK Intensity
Application.ActiveWorkbook.Worksheets(1).Cells(9, 10) = comboSiteClass.SelText ' Putting value of Scales Text

MsgBox "Scales TExt " & comboSiteClass.SelText & "Index " & comboSiteClass.ListIndex + 1


''''''''''''OUTPUT
directEcoLoss.Text = Application.ActiveWorkbook.Worksheets(1).Cells(11, 10)   'Direct economic Loss
totalDirectEcoLoss.Text = Application.ActiveWorkbook.Worksheets(1).Cells(12, 10)   ' total direct Eco Loss
noOfHomelessPeople.Text = Application.ActiveWorkbook.Worksheets(1).Cells(13, 10)  ' number of homeless people
dayTimePop.Text = Application.ActiveWorkbook.Worksheets(1).Cells(15, 8)  ' day time pop
nightTimePop.Text = Application.ActiveWorkbook.Worksheets(1).Cells(15, 11)   ' night time pop
dayTimeCasualties.Text = Application.ActiveWorkbook.Worksheets(1).Cells(16, 8)  ' day time casualties
nightTimeCasualties.Text = Application.ActiveWorkbook.Worksheets(1).Cells(16, 11) ' night time casualties
dayTimeInjuries.Text = Application.ActiveWorkbook.Worksheets(1).Cells(17, 8)  ' day time injuries
NightTimeInjuries.Text = Application.ActiveWorkbook.Worksheets(1).Cells(17, 11)  ' night time injuries

  Exit For
    ElseIf Application.CountA(Cells) = 0 Then
        'MsgBox ActiveSheet.Name & " is empty"
         
    End If
Next WS
    
    SourceWB.Close savechanges:=False
    Set WS = Nothing
    Set SourceWB = Nothing
    
WB.Activate
ASheet.Select
    ' Set ASheet = Nothing
    'Set WB = Nothing
    'Application.Quit
    
End Sub

Private Function getUnique(arrayMBT As Variant, arrCollect As Variant) As Variant

  Dim arr As New Collection, a
  Dim i As Long
  Dim count As Integer
  
  On Error Resume Next
  
  For Each a In arrayMBT
     arr.Add a, a
  Next

  For i = 1 To arr.count
'  'MsgBox "unique values are :: " & arr(i)
  Next
 ' 'MsgBox "arr count : " & arr.count
  For Z = 1 To Application.CountA(arrayMBT)
  ''MsgBox "arrMBT value s " & arrayMBT(Z)
  Next
 'count = arr.count
  'getUnique = arr
 
 For m = 0 To arr.count
  For n = 0 To Application.CountA(arrayMBT)
  
   If arr(m) = arrayMBT(n) Then
   'MsgBox "match from : " & arr(m) & " match to: " & arrayMBT(n) & " match at : " & n
   End If
  Next
 Next
 
''MsgBox "arrcollect size in another funciton" & Application.CountA(arrCollect)
 
 getUnique = arr
End Function

Private Sub ComboBoxSiteClass_Change()
With com

End Sub

Private Sub CommandButton1_Click()

   
    On Error GoTo 1
     'ActiveWorkbook.FollowHyperlink "c:\Windows\System32\Cleanmgr.exe", _
     'NewWindow:=True      '< for WinXP
    ActiveWorkbook.FollowHyperlink "F:\Survey\Cylinders_Nainital_Aligned.kmz", NewWindow:=True
    Exit Sub
1:           MsgBox Err.Description




End Sub

Private Sub CommandButton2_Click()

'Dim TheSheet As Worksheet
'Set TheSheet = Sheets("LookUpList")
MsgBox "count of Scales" & Application.CountA(Range("ER:ER"))
For i = 1 To Application.CountA(Range("ER:ER"))
If Len(Range("ER" & i)) > 0 Then
comboSiteClass.AddItem (Range("ER" & i))
End If

Next


End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub Label13_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub UserForm_Click()

End Sub




