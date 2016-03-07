Attribute VB_Name = "Module1"
Option Explicit

Public Const ToolBarName As String = "MyToolbarName"

'Custom data type for undoing
    Type SaveRange
        Val As Variant
        Addr As String
    End Type
    
'   Stores info about current selection
    Public OldWorkbook As Workbook
    Public OldSheet As Worksheet
    Public OldSelection() As SaveRange
    Dim MyRange As Range
    
    


'===========================================
Sub Auto_Open()
    Call CreateMenubar
End Sub

'===========================================
Sub Auto_Close()
    Call RemoveMenubar
End Sub

'===========================================
Sub RemoveMenubar()
    On Error Resume Next
    Application.CommandBars(ToolBarName).Delete
    On Error GoTo 0
End Sub

'===========================================
Sub CreateMenubar()

    Call RemoveMenubar

    Dim myCB As CommandBar
    Dim myCPup1 As CommandBarPopup
    Dim myCP1Btn1 As CommandBarButton
    Dim myCP1Btn2 As CommandBarButton
    Dim myCP1Btn3 As CommandBarButton
    Dim myCP1Btn4 As CommandBarButton
    Dim myCP1Btn5 As CommandBarButton
    Dim myCP1Btn6 As CommandBarButton
    Dim myCP1Btn7 As CommandBarButton
    Dim entry1, entry2 As String
    entry1 = "regular"
    entry2 = "reverse"

    
    ' Create a new Command Bar
    Set myCB = CommandBars.Add(Name:=ToolBarName, Position:=msoBarFloating)
    

    ' Add popup menu 1 to this bar
    Set myCPup1 = myCB.Controls.Add(Type:=msoControlPopup)
    With myCPup1
        .Caption = "MapIt!"
    End With
    
    ' Add button 1 to popup menu 1
    Set myCP1Btn1 = myCPup1.Controls.Add(Type:=msoControlButton)
    With myCP1Btn1
     .Caption = "DayGPO -> FCC"
     .Style = msoButtonIconAndCaption
     .FaceId = 1591
     .OnAction = "MapIt"
     .Parameter = entry1
    End With
    
    ' Add button 2 to popup menu 1
    Set myCP1Btn2 = myCPup1.Controls.Add(Type:=msoControlButton)
    With myCP1Btn2
     .Caption = "FCC -> DayGPO"
     .Style = msoButtonIconAndCaption
     .FaceId = 1590
     .OnAction = "MapIt"
     .Parameter = entry2
    End With
    
     
    ' Add button 3 to popup menu 1
    Set myCP1Btn3 = myCPup1.Controls.Add(Type:=msoControlButton)
    With myCP1Btn3
     .Caption = "Undo"
     .Style = msoButtonIconAndCaption
     .FaceId = 128
     .OnAction = "Undo"
    End With
    
    ' Add button 6 to popup menu 1
    Set myCP1Btn6 = myCPup1.Controls.Add(Type:=msoControlButton)
    With myCP1Btn6
     .Caption = "Options"
     .Style = msoButtonIconAndCaption
     .FaceId = 2933
     .OnAction = "ShowFrm"
    End With
    
    ' Add button 7 to popup menu 1
    Set myCP1Btn7 = myCPup1.Controls.Add(Type:=msoControlButton)
    With myCP1Btn7
     .Caption = "Eva File"
     .Style = msoButtonIconAndCaption
     .FaceId = 23
     .OnAction = "OpenFile"
    End With
    
    ' Add button 4 to popup menu 1
    Set myCP1Btn4 = myCPup1.Controls.Add(Type:=msoControlButton)
    With myCP1Btn4
     .Caption = "Help"
     .Style = msoButtonCaption
     '.FaceId = 176
     .OnAction = "NavigateToURL"
    End With
    
    ' Add button 5 to popup menu 1
    Set myCP1Btn5 = myCPup1.Controls.Add(Type:=msoControlButton)
    With myCP1Btn5
     .Caption = "About"
     .Style = msoButtonCaption
     .OnAction = "About"
    End With
      
    ' Show the command bar
    myCB.Visible = True
    

End Sub
Sub ShowFrm()
    
    OptionsFrm.Show
    
End Sub
Sub OpenFile()

    Dim MyPathFile As String
    Dim wb As Workbook

    
    On Err GoTo ErrHndlr
        MyPathFile = Workbooks("MapIt!.xlam").Path & "\DayGPOFCC.xlsx" 'Dictionary file
        Set wb = Workbooks.Open(MyPathFile)
    
ErrHndlr:

End Sub
Sub About()

    MsgBox "MapIt! " & Chr(169) & " 2013. All Rights Reserved.", vbOKOnly, "MapIt!"
    
End Sub


Sub MapIt()

Dim ctlCBarControl  As CommandBarControl
Dim entry As String
Dim iu As Integer
Dim c As Variant



    Set ctlCBarControl = CommandBars.ActionControl
    If ctlCBarControl Is Nothing Then Exit Sub
    'Examine the Parameter property of the ActionControl to determine
    'which control has been clicked
    entry = ctlCBarControl.Parameter
'MsgBox (entry)

'   Inserts zero into all selected cells


    Set MyRange = Range("A1:AA100")
'   The next block of statements
'   Save the current values for undoing
   
    ReDim OldSelection(MyRange.Count)

    Set OldWorkbook = ActiveWorkbook
    Set OldSheet = ActiveSheet
    iu = 0
    For Each c In MyRange
        iu = iu + 1
        OldSelection(iu).Addr = c.Address
        OldSelection(iu).Val = c.Formula
    Next c
            
'   Insert 0 into current selection
    Application.ScreenUpdating = False
    
'   Specify the Undo Sub
    Application.OnUndo "Undo the ZeroRange macro", "Undo"

Dim DayGPOrm As Variant
Dim FCCad As Variant
Dim PLacct As Variant
Dim Aacct As Variant
Dim DayGPO As Variant
Dim FCC As Variant
Dim Holder As Variant


Dim CSV As Variant
Dim FillArrayOld() As Variant
Dim FillArrayNew() As Variant

Dim wb As Workbook

Dim MyPathFile As String

Dim n As Long, j As Long, i As Long

'New Dimensions go here
Dim PLRowObj As New CSVPull
Dim HardwareObj As New CSVPull
Dim LabelsObj As New CSVPull
Dim ScenarioObj As New CSVPull



DayGPOrm = Array("All Management Reporting", "Publisher", "License Partner", "IP Ownership", "Frontline/Catalog")
FCCad = Array("Free to Play", "External Development Type", "ICO_Flag", " ", " ")

PLacct = Array("PL0000: P&L Accounts")
Aacct = Array("A000000: All Accounts")


'FillArray - CSV filenames to provide in function

PLRowObj.FillValues ("PLS")
HardwareObj.FillValues ("HardInt")
LabelsObj.FillValues ("Labels")
ScenarioObj.FillValues ("Scenario")

'Create Find/Replace Array
    
    DayGPO = Array(DayGPOrm, PLacct, PLRowObj.ReturnDayGPO, HardwareObj.ReturnDayGPO, LabelsObj.ReturnDayGPO, ScenarioObj.ReturnDayGPO)
    FCC = Array(FCCad, Aacct, PLRowObj.ReturnFCC, HardwareObj.ReturnFCC, LabelsObj.ReturnFCC, ScenarioObj.ReturnFCC)
    
'Check for Reverse

    If (entry = "reverse") Then
        Holder = DayGPO
        DayGPO = FCC
        FCC = Holder
    End If
    
'Check for CheckBox

    If OptionsFrm.CheckBox1.Value Then
        'Application.ReplaceFormat.Interior.ColorIndex = 36
    Else
        Application.ReplaceFormat.Interior.ColorIndex = 0
    End If
    
    For i = LBound(DayGPO) To UBound(DayGPO)

        For j = LBound(DayGPO(i)) To UBound(DayGPO(i))
            Cells.Replace What:=DayGPO(i)(j), Replacement:=FCC(i)(j), LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, ReplaceFormat:=True
        Next
    
    Next


'OpenWB

'Check for file show option
If OptionsFrm.CheckBox2.Value Then

    On Err GoTo ErrHndlr
        MyPathFile = Workbooks("MapIt!.xlam").Path & "\DayGPOFCC.xlsx" 'Dictionary file
        Set wb = Workbooks.Open(MyPathFile)
    
ErrHndlr:

End If

End Sub

Sub Undo()
'   Undoes the effect of the ZeroRange sub
    
Dim i As Integer

'   Tell user if a problem occurs
    On Error GoTo Problem

    Application.ScreenUpdating = False
    
'   Make sure the correct workbook and sheet are active
    OldWorkbook.Activate
    OldSheet.Activate
    
'   Restore the saved information
    For i = 1 To UBound(OldSelection)
        Range(OldSelection(i).Addr).Formula = OldSelection(i).Val
    Next i
    Exit Sub

'   Error handler
Problem:
    MsgBox "Nothing to Undo.", vbOKOnly, "MapIt!"
End Sub

Public Sub NavigateToURL()

  Const READYSTATE_COMPLETE As Integer = 4

  Dim objIE As Object
  
  Set objIE = CreateObject("InternetExplorer.Application")

  With objIE
    .Visible = True
    .Silent = True
   
    .Navigate "https://foo.bar" 'Help page URL
    Do Until .ReadyState = READYSTATE_COMPLETE
      DoEvents
    Loop
  End With
  
End Sub




