VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTestSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_lastErrorMessageCell As String = "C4"

''' <summary>   Gets the Last Error Message. </summary>
Friend Property Get LastErrorMessage() As String
    LastErrorMessage = Range(m_lastErrorMessageCell).value
End Property

''' <summary>   Sets the Last Error Message. </summary>
Friend Property Let LastErrorMessage(ByVal a_value As String)
    Range(m_lastErrorMessageCell).value = a_value
End Property

''' <summary>   Populates a combo box with the list of tests. </summary>
Private Sub ListTestsButton_Click()
    
    Dim p_moduleInfo As ModuleInfo
    testNamesCombo.Clear
    For Each p_moduleInfo In WorkbookUtilities.EnumerateProjectModules(cc_isr_Winsock_Test.ThisWorkbook.VBProject)
        DoEvents
        testNamesCombo.AddItem p_moduleInfo.FullName
    Next p_moduleInfo

End Sub

''' <summary>   Runs the selected test. </summary>
Private Sub RunTestsButton_Click()
    cc_isr_Test_Fx.TestExecutive.Execute UnitTestSheet, True
End Sub

