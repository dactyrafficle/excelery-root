Option Explicit

Public Sub TO_SQL_INSERT_INTO(control As IRibbonControl)

    Call WB_CONTENTS_AS_SQL_INSERT_INTO_BY_ROW

End Sub


Sub WB_CONTENTS_AS_SQL_INSERT_INTO_BY_ROW()

 'Dim table_name As String
 'table_name = Application.InputBox("TABLE NAME :")

 Dim wb_output As String
 wb_output = ""

 Dim i As Long
 For i = 1 To Worksheets.count

   Worksheets(i).Activate

   Dim ws_output As String
   ws_output = RANGE_TO_SQL_INSERT_INTO_STRING("COMPETITION_RESULTS2")
   wb_output = wb_output & ws_output

 Next i

 EXPORT_STR_AS_TXT_OPEN_IN_CHROME (wb_output)

End Sub