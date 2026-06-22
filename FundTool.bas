Attribute VB_Name = "FundTool"
' Fund & Portfolio Tool - button macros.
' This is the ONLY VBA in the project: three one-line subs, each calling a
' Python function via xlwings RunPython. All real logic lives in app.py.
'
' Requires: xlwings add-in installed, and in the VBA editor
' Tools > References > "xlwings" checked (so RunPython is available).

Option Explicit

Sub Btn_RecordFund()
    RunPython "import app; app.record_fund()"
End Sub

Sub Btn_Compute()
    RunPython "import app; app.compute_portfolio()"
End Sub

Sub Btn_Charts()
    RunPython "import app; app.make_charts()"
End Sub

Sub Btn_Evolution()
    RunPython "import app; app.evolution()"
End Sub
