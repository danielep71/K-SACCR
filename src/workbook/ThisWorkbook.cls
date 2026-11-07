VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit


Private Sub Workbook_Activate()
    'Refresh the K-SACCR tab for this workbook's state (#216).
        M_Ribbon.RefreshRibbon

End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    'Refresh the K-SACCR tab when a sheet of this workbook is activated (#216).
        M_Ribbon.RefreshRibbon

End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    'An input edit recalculates Results C5: refresh the tab once Excel is idle (#216).
        M_Main.RibbonStateChanged

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    'Keep the Ribbon's pointer out of the saved file (#216), and save the
    'navigation bar shown even while it is hidden (#620).
        M_Ribbon.HidePointer
        M_Main.NavigationBarBeforeSave

End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    'Store the Ribbon's pointer again for this session (#216), and hide the
    'navigation bar again if it was hidden before the save (#620).
        M_Ribbon.ShowPointer
        M_Main.NavigationBarAfterSave

End Sub
