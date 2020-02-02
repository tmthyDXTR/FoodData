Private Sub Form_Open(Cancel As Integer)

' Reset Nutrient preview list memory array
ReDim Preserve fdc_id_array(0)
fdc_id_array(0) = 12345

End Sub


Private Sub SearchBox_Change()
    
	' Typing in Search Box updates the Search List
    SearchList.Requery    

End Sub

Private Sub ResetButton_Click()
    
    ' Clear all Tables
    Call Reset_Table("Food_Basket")
    Call Reset_Table("Food_Basket_Nutrients")
          
    ' Reset Nutrient preview list memory array
    ReDim Preserve fdc_id_array(0)
    fdc_id_array(0) = 12345
    
    FoodBasket.Requery
    NutrientsPreviewList.Requery
    TotalNutrientsList.Requery
    
    Id_box.Value = Empty
    Name_box.Value = Empty
    Amount_box.Value = Empty
    
    MsgBox "Reset Tables"
    
End Sub


Private Sub SearchList_DblClick(Cancel As Integer)
    
    ' Set public nutrient key
    pub_fdc_id = SearchList.Column(1)
    'NutrientsPreviewList.Requery
    
    Id_box.Value = pub_fdc_id
    Name_box.Value = SearchList.Column(0)          
   
   ' Check if ingridient already in nutrient list
   ' if it is not, add it.
   If Check_fdc_added(pub_fdc_id, fdc_id_array) = False Then
    
        ' Resize fdc_id_array and add the new id
        ReDim Preserve fdc_id_array(UBound(fdc_id_array) + 1)
        fdc_id_array(UBound(fdc_id_array)) = pub_fdc_id
        
   
        Set mydb = CurrentDb()
        Set rs = mydb.OpenRecordset("SearchNutrients")
        Set nutrient_table = mydb.OpenRecordset("Food_Basket_Nutrients")
    
        ' Get row amount of nutrient table for new index
        rowAmount = nutrient_table.RecordCount
        n_index = rowAmount
    
            
        ' Copy from nutrient search query to nutrient table
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst 'Unnecessary in this case, but still a good habit
            Do Until rs.EOF = True
            
                If rs![amount] > 0 Then
                
                    nutrient_table.AddNew
                    n_index = n_index + 1
                    nutrient_table![id] = n_index
                    nutrient_table![fdc_id] = rs![fdc_id]
                    nutrient_table![amount] = rs![amount]
                    nutrient_table![actual_amount] = rs![amount]
                    nutrient_table![Name] = rs![Name]
                    nutrient_table![unit_name] = rs![unit_name]
                    nutrient_table.Update
                End If
                rs.MoveNext
                
            Loop
        End If
        
        
        nutrient_table.Close
        rs.Close
        mydb.Close
        
        Set rs = Nothing
        Set mydb = Nothing
        Set nutrient_table = Nothing
   End If
   NutrientsPreviewList.Requery

End Sub

   
Private Sub FoodBasket_DblClick(Cancel As Integer)

    ' Select fdc_id
    pub_fdc_id = FoodBasket.Column(0)
    

    ' Update Nutrient Preview List
    NutrientsPreviewList.Requery
    
    
    ' RÅckwÑrts Selektion
    Id_box.Value = FoodBasket.Column(0)
    Name_box.Value = FoodBasket.Column(1)
    Amount_box.Value = FoodBasket.Column(2)     

End Sub


Private Sub Add_Button_Food_Basket_Click()
    
    item_name = Name_box.Value
    item_id = Id_box.Value
    item_amount = Amount_box.Value
    
    'Dim UniqueFlag As Boolean
    'UniqueFlag = False
    
    pub_fdc_id = item_id
    
    
    ' open the table
    Set mydb = CurrentDb()
    Set rst = mydb.OpenRecordset("Food_Basket")
    
    'DebugText.Value = Check_fdc_added_food_basket(pub_fdc_id, rst, item_amount)
    
    If Check_fdc_added_food_basket(pub_fdc_id, rst, item_amount) = False Then
        rst.AddNew
        rst![fdc_id] = item_id
        rst![Name] = item_name
        rst![amount] = item_amount
        rst.Update
        
        ' Loop through Food Basket Nutrient list and edit actual amounts
        Set nutrient_table = mydb.OpenRecordset("Food_Basket_Nutrients")
        If Not (nutrient_table.EOF And nutrient_table.BOF) Then
            nutrient_table.MoveFirst 'Unnecessary in this case, but still a good habit
            Do Until nutrient_table.EOF = True
            
                If nutrient_table![fdc_id] = pub_fdc_id Then
                
                    nutrient_table.Edit
                    nutrient_table![actual_amount] = (nutrient_table![amount] / 100) * Amount_box.Value
                    nutrient_table.Update
                End If
                nutrient_table.MoveNext
            Loop
        End If
        nutrient_table.Close
        Set nutrient_table = Nothing  'Clean up
        
    Else
        ' Loop through Food Basket Nutrient list and edit actual amounts
        Set nutrient_table = mydb.OpenRecordset("Food_Basket_Nutrients")
        If Not (nutrient_table.EOF And nutrient_table.BOF) Then
            nutrient_table.MoveFirst 'Unnecessary in this case, but still a good habit
            Do Until nutrient_table.EOF = True
            
                If nutrient_table![fdc_id] = pub_fdc_id Then
                
                    nutrient_table.Edit
                    nutrient_table![actual_amount] = (nutrient_table![amount] / 100) * Amount_box.Value
                    nutrient_table.Update
                End If
                nutrient_table.MoveNext
            Loop
        End If
        nutrient_table.Close
        Set nutrient_table = Nothing  'Clean up
        
        MsgBox "Nutrient amount updated."
        
    End If
    
    rst.Close 'Close the recordset
    mydb.Close
    
    Set rst = Nothing
    Set mydb = Nothing
    
    
    NutrientsPreviewList.Requery
    FoodBasket.Requery
    TotalNutrientsList.Requery
    
End Sub


Private Sub Delete_Button_Food_Basket_Click()

    Call Remove_Ingredient(pub_fdc_id, "Food_Basket")
    
    FoodBasket.Requery
    TotalNutrientsList.Requery
    
    MsgBox "Ingredient removed."

End Sub