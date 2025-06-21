Attribute VB_Name = "Scene"
Sub DelEmptyCol(control As IRibbonControl)
Call ModulesAddins.DelEmptyColMod
End Sub
Sub DelEmptyRow(control As IRibbonControl)
Call ModulesAddins.DelEmptyRowMod
End Sub
Sub LvlUnload(control As IRibbonControl)
Call ModulesAddins.LvlUnloadMod
End Sub
Sub LvlPivot(control As IRibbonControl)
Call ModulesAddins.LvlPivotMod
End Sub
Sub Summary(control As IRibbonControl)
Call ModulesAddins.SummaryMod
End Sub
Sub Format(control As IRibbonControl)
Call LvlFormat.FormatMod
End Sub
Sub Group(control As IRibbonControl)
Call ModulesAddins.GroupMod
End Sub
Sub UnGroup(control As IRibbonControl)
Call ModulesAddins.UnGroupMod
End Sub
Sub DownOpen(control As IRibbonControl)
Call ModulesAddins.DownOpenMod
End Sub
Sub RightOpen(control As IRibbonControl)
Call ModulesAddins.RightOpenMod
End Sub
Sub Data(control As IRibbonControl)
Call ModulesAddins.Форматирование_дат
End Sub
Sub SmartInj(control As IRibbonControl)
Call Smartart.BuildSmartArtFromPivot
End Sub
Sub ValueIdentLvl(control As IRibbonControl)
Call Smartart.SetIndentByValue
End Sub
