VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Created by James Innes
'The below should record changes to cells into Columns A & B of the associated row, by storing selected cells into a dictionary and on worksheet change comparing the two values.
'Caveats: This will not work of cells which automatically update (i.e. via a formula). After prolonged use the dictionary will get pretty large, and sloooowwwww dooooowwwwn.

Dim OldVals As New Dictionary

Dim ShouldRecord As Boolean

Const COL_RANGE = 25 'Column range for logging values (Column AA)
Const ROW_RANGE = 15000 'Row range for logging values
Const MAX_CELLS = 1000 'Max number of cells to log; when over MAX_CELLS are selected they will not be stored in the dictionary

'Function to log selected cells into a dictionary (OldVals), in order to track changes

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim newTarget As Range
    Set newTarget = Target
    Dim cell As Range
    Dim cellCount As Long
    
    ShouldRecord = True
    
    If newTarget.CountLarge > MAX_CELLS Then 'Check if original selection is over MAX_CELLS allowance for logging
        If newTarget.Columns.Count > COL_RANGE Then 'Remove columns (Past column AA) from the selection, since they are not used
            Set newTarget = newTarget.Resize(newTarget.Rows.Count, COL_RANGE)
        End If
        If newTarget.Rows.Count > ROW_RANGE Then 'Remove rows (past row 15000) from selection, since they are not used
            Set newTarget = newTarget.Resize(ROW_RANGE, newTarget.Columns.Count)
        End If
        If newTarget.CountLarge > MAX_CELLS Then 'If even after the selection shrinking the number of cells selected is larger than MAX_CELLS, they won't be logged
            MsgBox ("Warning: Too many cells have been selected to record spread sheet changes. Any updates made to this selection will not be recorded.")
            ShouldRecord = False
        End If
    End If
    
    If ShouldRecord Then
        For Each cell In newTarget 'Loop through each selected cell, and add then to the OldVals dictionary
            If OldVals.Exists(cell.Address) Then
                Debug.Print "New value of " & cell.Address & " is " & CStr(cell.Value) & "; old value was " & OldVals(cell.Address)
            Else
                Debug.Print "No old value for " + cell.Address
            End If
            OldVals(cell.Address) = CStr(cell.Value)
        Next
        Debug.Print "Number of recorded cells: " & newTarget.Count
    End If
    
    Target.Calculate
    
End Sub

'Function will check new cell value against logged value (OldVals), and record any changes in column A

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim OldValueString As String
    Dim NewValueString As String

    If ShouldRecord Then
        If Target.Column > 2 Then 'Check to prevent script looping on itself due to changing cell values
            For Each cell In Target 'Loop through all of the selected cells
                OldValueString = OldVals(cell.Address) 'Variable for the old contents of the cell, stored in our dictionary
                NewValueString = CStr(cell.Value) 'Variable for the new contents of the cell, needs to be converted to a string
                If cell.Column > COL_RANGE Or cell.Row > ROW_RANGE Then 'Check to make sure that the
                    MsgBox (cell.Address & " is outwith the cell value logging range. Updates will NOT be recorded.")
                Else
                    If NewValueString <> OldValueString Then
                        Cells(cell.Row, 1).Value = Now
                        If Cells(cell.Row, 2).Value <> "" Then
                            Cells(cell.Row, 2).Value = Now & " - " & cell.Address & " Change: " & OldValueString & " >> " & NewValueString & Chr(10) & Cells(cell.Row, 2).Value
                        Else
                            Cells(cell.Row, 2).Value = Now & " - " & cell.Address & " Change: " & OldValueString & " >> " & NewValueString
                        End If
                    End If
                End If
            Next
        End If
    End If
End Sub
