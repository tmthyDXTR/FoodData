Option Compare Database

Public pub_fdc_id As String

Public n_index As Integer




Public fdc_id_array() As String



Option Explicit






Public Function Get_fdc_id() As String

  Get_fdc_id = pub_fdc_id

End Function


Public Function Check_fdc_added(fdc_id As Variant, arr As Variant) As Boolean
'DESCRIPTION: Function to check if ingridient already added to nutrient list
'INPUT: Pass the function a fdc_id to search for and an array of fdc_ids of any data type.
'OUTPUT: True if is in array, false otherwise
Dim element As Variant
On Error GoTo Check_fdc_added_Error: 'array is empty
    For Each element In arr
        If element = fdc_id Then
            Check_fdc_added = True
            Exit Function
        End If
    Next element
Exit Function
Check_fdc_added_Error:
On Error GoTo 0
Check_fdc_added = False
End Function


Public Function Check_fdc_added_food_basket(fdc_id As Variant, rs As Variant, ByVal new_amount As String) As Integer


    
    'Check to see if the recordset actually contains rows
    If Not (rs.EOF And rs.BOF) Then
    'MsgBox "hallo"
        rs.MoveFirst 'Unnecessary in this case, but still a good habit
        Do Until rs.EOF = True
        'MsgBox "hallohallo"
            If CStr(rs![fdc_id]) = fdc_id Then
                Check_fdc_added_food_basket = True
                rs.Edit
                rs![amount] = new_amount
                rs.Update
                'MsgBox "hallohallohallo"
                Exit Function
            End If
            'Move to the next record. Don't ever forget to do this.
            rs.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If
    
    Exit Function
    Check_fdc_added_food_basket = False



End Function


Public Function ResetTable(table As String)
    
    Dim dbs As Variant, rst As Recordset
    
    Set dbs = CurrentDb()

   ' Delete
    dbs.Execute "DELETE * FROM " _
       & ""&table&";"
    
    dbs.Close

End Function




