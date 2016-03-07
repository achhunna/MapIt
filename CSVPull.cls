VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVPull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DayGPO() As Variant

Private FCC() As Variant

Dim MyPathFile As String
Dim MyData As String, strData() As String
Dim n As Long, j As Long, i As Long

Dim FinalArray() As String


Public Sub FillValues(CSV As String)

    MyPathFile = Workbooks("MapIt!.xlam").Path & "\CSV\" & CSV & ".csv"
    
    Open MyPathFile For Binary As #1
    MyData = Space$(LOF(1))
    Get #1, , MyData
    Close #1
    
    strData() = Split(MyData, vbCrLf)
    
    ReDim Preserve FinalArray(j + UBound(strData) + 1)
    
    For i = LBound(strData) To UBound(strData) - 1
        FinalArray() = Split(strData(i), ",")
        
        ReDim Preserve DayGPO(i)
        ReDim Preserve FCC(i)
        DayGPO(i) = FinalArray(0)
        FCC(i) = FinalArray(1)
        
    Next


End Sub

Public Property Get ReturnDayGPO() As Variant

    ReturnDayGPO = DayGPO()
    
End Property

Public Property Get ReturnFCC() As Variant

    ReturnFCC = FCC()
    
End Property


