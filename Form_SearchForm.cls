VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Open(Cancel As Integer)



ReDim Preserve fdc_id_array(0)
fdc_id_array(0) = "12345"



End Sub


Private Sub SearchBox_Enter()
    
    SearchList.Requery
    

End Sub

Private Sub SearchList_DblClick(Cancel As Integer)
    
    Id_box.Value = SearchList.Column(1)
    Name_box.Value = SearchList.Column(0)
    
    
    
    ' Set public nutrient key
    pub_fdc_id = Id_box.Value
    'NutrientsPreviewList.Requery
   
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
            
                nutrient_table.AddNew
                n_index = n_index + 1
                nutrient_table![id] = n_index
                nutrient_table![fdc_id] = rs![fdc_id]
                nutrient_table![amount] = rs![amount]
                nutrient_table![Name] = rs![Name]
                nutrient_table![unit_name] = rs![unit_name]
                nutrient_table.Update
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
    
    
    ' Rückwärts Selektion
    Id_box.Value = FoodBasket.Column(0)
    Name_box.Value = FoodBasket.Column(1)
    Amount_box.Value = FoodBasket.Column(2)
    
    


End Sub

Private Sub Addbutton_to_basket_Click()
    
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
        
        
    Else
        
        
    End If
    
    rst.Close 'Close the recordset
    mydb.Close
    Set rst = Nothing 'Clean up
    Set mydb = Nothing
    
 
    
    FoodBasket.Requery
    
    
End Sub
