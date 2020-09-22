Attribute VB_Name = "LVAscDec"
Public Function EnhListView_SortColumns( _
                lstListViewName As ListView, _
                usdColIndex, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '_______________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_SortColumns
    
    '_______________________________________________________________________
    ' set function return to true
    EnhListView_SortColumns = True
    
    '_______________________________________________________________________
    ' if there are columns to go through...
    If lstListViewName.ListItems.Count > 0 Then
        ' if the sort property is turned off turn it on
        If lstListViewName.Sorted = False Then lstListViewName.Sorted = True
        ' set the sortby column
        lstListViewName.SortKey = _
            lstListViewName.ColumnHeaders.Item(usdColIndex).Index - 1
        ' if it's sorted ascending
        If lstListViewName.SortOrder = lvwAscending Then
            ' sort it descending
            lstListViewName.SortOrder = lvwDescending
        ' if it's sorted descending
        Else
            ' sort it ascending
            lstListViewName.SortOrder = lvwAscending
        End If
    End If
    
    '_______________________________________________________________________
    ' exit before error handler
    Exit Function
    
'_______________________________________________________________________
' deal with errors
err_EnhListView_SortColumns:
    
    '_______________________________________________________________________
    ' set function return to false
    EnhListView_SortColumns = False
    '_______________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_SortColumns"
    End If
    
    '_______________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_SortColumns" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    'Debug.Assert False
    
    '_______________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

