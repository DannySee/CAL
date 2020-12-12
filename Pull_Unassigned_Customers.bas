Attribute VB_Name = "Format"


Sub Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Update listbox
    Utility.UpdateListbox(False)

    'Show utility elements
    Utility.Show(oBtnPull.GetShapes)

End Sub


Sub SelectCst()

End Sub


Sub Cancel()

End Sub
