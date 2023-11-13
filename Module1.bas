Attribute VB_Name = "Module1"
Option Explicit

'Macros() - Opens 1st UserForm with list of available macros with option to delete or editing existing macro and create a new one
'       If a macro is selected, calls get_fields() to get list of fields for the pivot table selected,
'       then passes it as a dictionary to the function report()
'       If the checkbox for macro editing/creating is selected, it calls edit_personal(), after which it will exit back to this sub

'report() - Calls sheet_name_generator() to create a non existing sheet name.
'       Then calls create_pivot() to add a worksheet (if needed) and to create a pivot table with the fields in the personal workbook 
'       (passed with the dictionary "fields")

Dim wkb As Workbook, a_wkb As Workbook, headers as range, Dic As Object, opt as Object, form2 As New pe_form, levels1 as Scripting.dictionary, levels2 as Scripting.dictionary, cmd_dict as Scripting.dictionary

Sub Macros()
' Keyboard Shortcut: Ctrl+Shift+M
refresh:

    Set a_wkb = ActiveWorkbook
    Set wkb = Application.Workbooks("PERSONAL.xlsb")

    'Initialize form
    Dim form As New pe_form
    Dim opt As Object
    Dim chkbox1 As Object
    Dim title As Object
    Dim radio As Control
    Dim key As Variant, i
    Dim macros_list As String: macros_list = ""
    Set form = Nothing

    'Create buttons in form
    'Get list of macros

    Dim ws As Worksheet
    i = 1
    For Each ws In wkb.Worksheets
        macros_list = ws.name & ";" & macros_list
    Next ws

    For Each key In Split(macros_list, ";")
        If Len(key) > 0 Then
            Set opt = form.Controls.add("Forms.OptionButton.1", "radioBtn" & i, True)
            With opt
                .Caption = key
                .Top = opt.Height * i
                .GroupName = "Macros"
                .Width = 300
            End With
            i = i + 1
        End If
    Next
    
    'Checkbox for edit macro
    Set chkbox1 = form.Controls.add("Forms.CheckBox.1", "checkbox1", True)
    With chkbox1
            .Caption = "Edit Pivot Macro"
            .Top = opt.Height * i
            .GroupName = "Macros"
            .Width = 300
    End With
    i = i + 1

    'Checkbox for delete
    Dim chkbox3 As Object: Set chkbox3 = form.Controls.add("Forms.CheckBox.1", "checkbox3", True)
    With chkbox3
            .Caption = "Delete Pivot Macro"
            .Top = opt.Top + chkbox1.Height + 20
            .GroupName = "delete"
            .Width = 300
    End With

    'Checkbox to add macro
    Dim chkbox4 As Object: Set chkbox4 = form.Controls.add("Forms.CheckBox.1", "checkbox4", True)
    With chkbox4
            .Caption = "Add Pivot Macro"
            .Top = opt.Top + chkbox1.Height * 2 + 20
            .GroupName = "add"
            .Width = 150
    End With

    Dim textbox1 As Object: Set textbox1 = form.Controls.add("Forms.TextBox.1", "textbox1", True)
    With textbox1
            .Top = opt.Top + chkbox1.Height * 2 + 20
            .Width = 150
            .left = chkbox4.width
    End With    

    form.CommandButton1.Top = opt.Top + chkbox1.Height * 3 + 20
    form.Width = chkbox4.width + textbox1.width + 50
    form.Height = opt.Top + chkbox1.Height * 3 + form.CommandButton1.Height + 50
    

select_report:
    'Show form
    form.Show 'vbModal
    
    dim name as string: name = ""
    dim macro as string: macro = ""
    if chkbox4 = True then
        
        if textbox1.value = vbNullString then goto refresh

        Dim headers(4) As String
        headers(0) = "Filters"
        headers(1) = "Rows"
        headers(2) = "Columns"
        headers(3) = "Data"

        Dim new_ws As Worksheet: Set new_ws = wkb.Sheets.add
        name = sheet_name_generator(textbox1.value, wkb)
        new_ws.name = name
        new_ws.Range("A1:D1").value = headers
        
        macros_list = macros_list & name
       
        edit_personal macros_list, name
        Set form = Nothing
        GoTo refresh
    
    else 
        
        For Each radio In form.Controls
            If TypeName(radio) = "OptionButton" And radio = True Then
                macro = radio.caption
            end if
        next radio
        
        if chkbox1 = True Then
            edit_personal macros_list, macro
            Set form = Nothing
            GoTo refresh
        
        elseIf chkbox3 = True Then
            Application.DisplayAlerts = False
            wkb.Worksheets(macro).Delete
            Application.DisplayAlerts = True
            Set form = Nothing
            GoTo refresh
        
        End If
    end if    

    'Unload form
    Set form = Nothing
    
    'Get list of fields for pivot
    Dim fields_list As Scripting.Dictionary
    Set fields_list = New Scripting.Dictionary
    
    If macro = "DSV" Then
        Dim vUserInput As VbMsgBoxResult
        vUserInput = MsgBox("Run Query?" & vbCrLf & "If you click no, it will only run the pivot macro and not the Toad Query", vbYesNoCancel)
        Select Case vUserInput
        Case vbYes
            DSV
        Case vbCancel
            end
        End Select
    End If
    Set fields_list = get_fields(wkb, macro)    
    
    report macro, fields_list

End Sub

Function report(name As String, fields As Scripting.Dictionary)
    dim wkb as workbook: set wkb = ActiveWorkbook
    name = sheet_name_generator(name, wkb)
    create_pivot name
    
    Dim var As Variant: var = ActiveCell.PivotTable.name
    Dim pivot As PivotTable: Set pivot = ActiveSheet.PivotTables(var)

    'On Error Resume Next
    Dim item As Variant, i As Integer, fld As Variant
    
    For Each item In fields.Keys
        i = 1
        For Each fld In Split(fields(item), ";")
            fld = CStr(Replace(fld, ";", ""))
            Debug.Print Application.Match(fld, headers.value)
            If Not IsError(Application.Match(fld, headers, 0)) Then
                If Not fld = vbNullString Then
                    If item = "Sum of" Then
                        pivot.AddDataField ActiveSheet.PivotTables( _
                        var).PivotFields(fld), _
                        item & " " & fld, xlSum
                    Else
                        With pivot.PivotFields(fld)
                            .orientation = item
                            .position = i
                            i = i + 1
                        End With
                    End If
                End If
            End If
        Next
    Next

    'Change format of data fields
    For Each fld In Split(fields("Sum of"), ";")
        fld = CStr(Replace(fld, ";", ""))
        If Not fld = vbNullString Then
            With pivot.PivotFields( _
                "Sum of " & fld)
                .NumberFormat = "#,##0.00_);(#,##0.00)"
            End With
        End If
    Next
    
    If name = "DSV" Then Toad pivot
    'Application.ScreenUpdating = True

End Function

Function Toad(pivot As PivotTable)

    'Get pivot range for analysis
    Dim pivot_range As Range: Set pivot_range = pivot.GetPivotData()
    
    'Get first row of pivot
    Dim pivot_row As Integer: pivot_row = pivot.GetPivotData().CurrentRegion().row()
    
    'Get first column of pivot
    Dim pivot_col1 As Integer: pivot_col1 = pivot.GetPivotData().CurrentRegion().column()
    
    'Get last column of pivot
    Dim pivot_col As Integer: pivot_col = pivot.GetPivotData().column()
    
    'Get last row of pivot
    Dim pivot_last_row As Integer: pivot_last_row = pivot.GetPivotData().row()
    
    Dim quantity_col As Long: quantity_col = Cells(pivot_row, 1).EntireRow.Find(what:="QUANTITY", _
    LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).column
    
    Dim net_price As Long: net_price = Cells(pivot_row, 1).EntireRow.Find(what:="Sum of REPORTED_NET_UNIT_PRICE", _
    LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).column
    
    'Hide Quantity and Net Price columns
    Dim Net_Price_Col_letter As String: Net_Price_Col_letter = Split(Cells(1, net_price).Address(True, False), "$")(0)
    Dim quantity_col_letter As String: quantity_col_letter = Split(Cells(1, quantity_col).Address(True, False), "$")(0)
    
    Range(Net_Price_Col_letter & ":" & Net_Price_Col_letter).EntireColumn.Hidden = True
    Range(quantity_col_letter & ":" & quantity_col_letter).EntireColumn.Hidden = True
    
    'Insert bookings header
    Cells(pivot_row, pivot_col + 1).value = "Bookings"
    
    '(Loop) Get values of both columns for calculations
    Dim x As Integer
    For x = (pivot_row + 1) To (pivot_last_row - 1)
        
        Range(Cells(x, pivot_col + 1), Cells(x, pivot_col + 1)).Formula = "=if(isblank(" & Net_Price_Col_letter & x & "),""""," & Net_Price_Col_letter & x & " * " & quantity_col_letter & x & ")"
        Range(Cells(x, pivot_col + 1), Cells(x, pivot_col + 1)).NumberFormat = "[$$-chr-Cher-US]#,##0.00;[Red]-[$$-chr-Cher-US]#,##0.00"
        
    Next
    'Application.ScreenUpdating = True
End Function

Function transposeArray(myarr As Variant) As Variant
    
    Dim myvar As Variant
    On Error Resume Next
    myvar = [[]]
    ReDim myvar(LBound(myarr, 2) To UBound(myarr, 2), LBound(myarr, 1) To UBound(myarr, 1))
    Dim i, j As Integer
    For i = LBound(myarr, 2) To UBound(myarr, 2)
        For j = LBound(myarr, 1) To UBound(myarr, 1)
            myvar(i, j) = myarr(j, i)
        Next
    Next
    On Error GoTo 0
    transposeArray = myvar
End Function

Function DSV()

    set a_wkb = ActiveWorkbook
    If ActiveCell.value = "" Then GoTo no_data
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim mtxData As Variant
     
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    Dim order_type As String: order_type = InputBox("DSV, ERP or XASRV")
    If order_type = "" Then Exit Function
    
    On Error GoTo vpn
    
    If LCase(order_type) = "dsv" Or LCase(order_type) = "erp" Then
        cn.Open ( _
        "User ID= __ " & _
        ";Password= __ " & _
        ";Data Source= __ " & _
        ";Provider= __ ")
    ElseIf LCase(order_type) = "xasrv" Then
        cn.Open ( _
        "User ID= __ " & _
        ";Password= __ " & _
        ";Data Source= __ " & _
        ";Provider= __ ")
    End If
    
    rs.CursorType = adOpenForwardOnly
    
    Dim list_range As Range: Set list_range = Range(ActiveCell.Address, ActiveCell.End(xlDown).Address)
    Dim list As String: list = ""
    Dim i As Integer
    If ActiveCell.Offset(1, 0).value = "" Then
        list = "'" & list_range(1) & "'"
    Else
        For i = 1 To list_range.count
            list = list & "'" & list_range(i) & "'"
            If Not i = list_range.count Then
                list = list & ","
            End If
        Next
    End If
    Debug.Print list
    
    If LCase(order_type) = "dsv" Then
        Dim query As String: query = "select * " & _
            "from wips_bookings where 1 = 1 " & _
            "and (trans_id in (" & list & ") or pos_trans_id in (" & list & "))"
    ElseIf LCase(order_type) = "erp" Then
        query = "select * " & _
            "from wips_bookings where 1 = 1 " & _
            "and ERP_ORDER_NUMBER in (" & list & ")"
    ElseIf LCase(order_type) = "xasrv" Then
        query = "select HEADER_ID,ORDER_NUMBER from xxopl.xxopl_order_headers_all where 1 = 1 " & _
            "and header_id in (" & list & ")"
    Else
        MsgBox ("Type not recognized")
        Exit Function
    End If
    Debug.Print query
    rs.Open (query), cn
    
    On Error GoTo no_data:
    mtxData = rs.GetRows

    mtxData = transposeArray(mtxData)
    
    Dim sheet_name As String: sheet_name = order_type & "_"
    
    dim ws as worksheet: Set ws = Sheets.add(after:= _
        Sheets(Sheets.count))
        
    ws.name = sheet_name_generator(sheet_name, a_wkb)

    Dim iCols As Integer
    For iCols = 0 To rs.fields.count - 1
        ws.Cells(1, iCols + 1).value = rs.fields(iCols).name
    Next
    
    ws.Range("A2").Resize(UBound(mtxData, 1) + 1, UBound(mtxData, 2) + 1).value = mtxData
    
    'Cleanup in the end
    Set rs = Nothing
    Set cn = Nothing
    
    'Run Toad Pivot
    If LCase(order_type) = "xasrv" Then end
    exit function
vpn:
    MsgBox "VPN not connected"
    end

no_data:
    MsgBox "No data"
    end
    
End Function

Function sheet_name_generator(sheet_name As String, wkb As Workbook)

        Dim sheets_string As String, new_sheet_name As String
        
        Dim ws
        Dim n As Integer: n = 0
        new_sheet_name = sheet_name
        try_again:
        n = n + 1

        For Each ws In wkb.Worksheets
            if ws.name = new_sheet_name then 
                new_sheet_name = sheet_name & " (" & n & ")"
                goto try_again
            end if
        Next

        sheet_name_generator = new_sheet_name
        
End Function

Function create_pivot(name)

    'Application.ScreenUpdating = False
    On Error Resume Next
    Dim pivot_exists As String: pivot_exists = ActiveCell.PivotTable.name
    On Error GoTo 0
    If Not pivot_exists = vbNullString Then GoTo no_pivot:
    'Declare Variables
    Dim Pivot_Sheet As Worksheet
    Dim Data_Sheet As Worksheet
    Dim PCache As PivotTable
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim lastRow As Long
    Dim LastCol As Long
    
    'Define Data Range
    Set PRange = ActiveCell.CurrentRegion
    If Not PRange.Rows.count > 1 Or PRange.Cells(1, 1).Value = "" Then
        MsgBox "Data for pivot not selected"
        end
    end if

    set headers = PRange.Rows(1)

    'Insert a New Blank Worksheet
    Application.DisplayAlerts = False
    Set Data_Sheet = Worksheets(ActiveSheet.name)
    Sheets.add before:=ActiveSheet
    ActiveSheet.name = name
    Application.DisplayAlerts = True
    Set Pivot_Sheet = Worksheets(name)

    'Insert Pivot Table
    Set PCache = ActiveWorkbook.PivotCaches.Create _
    (SourceType:=xlDatabase, SourceData:=PRange).CreatePivotTable _
    (TableDestination:=Pivot_Sheet.Cells(1, 1), TableName:="PivotTable22")
    
    'Remove Pivot Classic View
    ActiveSheet.PivotTables("PivotTable22").InGridDropZones = False

no_pivot:
End Function

Function get_fields(wkb As Workbook, macro As String)

    Dim item As Variant
    Dim key As String
    Dim fields As Scripting.Dictionary
    Set fields = New Scripting.Dictionary

    Dim ws As Worksheet: Set ws = wkb.Worksheets(macro)
    
    Dim rng As Range, cell

    Dim filter_str As String
    Set rng = ws.Range("A2:" & ws.Range("A100").End(xlUp).Address)
    For Each cell In rng
        filter_str = filter_str & ";" & cell.value
    Next
    fields(xlPageField) = filter_str

    Dim row_str As String
    Set rng = ws.Range("B2:" & ws.Range("B100").End(xlUp).Address)
    For Each cell In rng
        row_str = row_str & ";" & cell.value
    Next
    fields(xlRowField) = row_str

    Dim col_str As String
    Set rng = ws.Range("C2:" & ws.Range("C100").End(xlUp).Address)
    For Each cell In rng
        col_str = col_str & ";" & cell.value
    Next
    fields(xlColumnField) = col_str

    Dim sum_str As String
    Set rng = ws.Range("D2:" & ws.Range("D100").End(xlUp).Address)
    For Each cell In rng
        sum_str = sum_str & ";" & cell.value
    Next
    fields("Sum of") = sum_str

    Set get_fields = fields

    Windows("PERSONAL.XLSB").Visible = False

End Function

Function edit_personal(ByRef macros_list As String, optional macro as string)
    
    Set form2 = Nothing

    Dim PRange As Range
    Dim headers(4) As String
    headers(0) = "Filters"
    headers(1) = "Rows"
    headers(2) = "Columns"
    headers(3) = "Data"

    Set opt = form2.Controls.add("Forms.ComboBox.1", "ComboBox", True)
    opt.Width = 200
    dim m
    for each m in Split(macros_list, ";")
        if m <> "" then opt.AddItem m
    next m
    opt.ListIndex = 0

    if not macro = vbNullString then opt.value = macro

    Dim comboArray() As New ComboClass
    ReDim Preserve comboArray(1 to 1)

    set comboArray(1).BoxSelect = opt

    dim label As Object, head as Variant
    Dim u As Integer: u = 1

    Dim cmdArray() As New ButtonClass
    ReDim Preserve cmdArray(1 To 4)

    Set levels1 = New Scripting.Dictionary
    Set levels2 = New Scripting.Dictionary
    Set cmd_dict = New Scripting.Dictionary

    For Each head In headers: Do

        If head = "" Then Exit Do
        
        Set label = form2.Controls.add("Forms.Label.1", "Label" & u, True)
        'Listbox Existing Fields (left listbox)
        levels1.add u, form2.Controls.add("Forms.ListBox.1", "ListBox" & u, True)
        
        'Listbox Fields for Pivot (right listbox)
        levels2.add u, form2.Controls.add("Forms.ListBox.1", "ListBox" & u & u, True)
        
        'Button Add
        cmd_dict.add u, form2.Controls.add("Forms.CommandButton.1", "add" & u, True)
        cmd_dict.add replace(space(2), " ", CStr(u)), form2.Controls.add("Forms.CommandButton.1", "up" & u, True)
        cmd_dict.add replace(space(3), " ", CStr(u)), form2.Controls.add("Forms.CommandButton.1", "down" & u, True)
        cmd_dict.add replace(space(4), " ", CStr(u)), form2.Controls.add("Forms.CommandButton.1", "delete" & u, True)

        With label
            .Caption = head
            .Width = 30
            .left = 5
            If u = 1 Then
                .Top = opt.Height + 5
            Else
                .Top = levels1(u-1).top + levels1(u-1).height + 20
            End If
        End With
        
        With levels1(u)
            .Height = 70
            .left = label.Width + 40
            .Width = 150
            .MultiSelect = 1
            .ListStyle = 0
            If u = 1 Then
                .Top = opt.Height + 10
            Else
                .Top = levels1(u-1).top + .height + 20
            End If
        End With

        With levels2(u)
            .Height = 70
            .left = levels1(u).left + levels1(u).Width + 70
            .Width = 150
            .MultiSelect = 1
            .top = levels1(u).top
        End With

        with cmd_dict(u)
            .width = 30
            .Height = 20
            .left = levels1(u).left + levels1(u).width + 20
            .caption = "Add"
            .Top = levels1(u).top
        end with

        with cmd_dict( replace(space(2), " ", CStr(u)) )
            .width = 30
            .Height = 20
            .left = levels1(u).left + levels1(u).width + 20
            .caption = "Up"
            .Top = cmd_dict(u).top + .height
        end with

        with cmd_dict( replace(space(3), " ", CStr(u)) )
            .width = 30
            .Height = 20
            .left = levels1(u).left + levels1(u).width + 20
            .caption = "Down"
            .Top = cmd_dict(u).top + (.height * 2)
        end with

        with cmd_dict( replace(space(4), " ", CStr(u)) )
            .width = 30
            .Height = 20
            .left = levels1(u).left + levels1(u).width + 20
            .caption = "Del"
            .Top = cmd_dict(u).top + (.height* 3)
        end with

        Set cmdArray(u).CmdAdd = cmd_dict(u)
        Set cmdArray(u).CmdUp = cmd_dict( replace(space(2), " ", CStr(u)) )
        Set cmdArray(u).CmdDown = cmd_dict( replace(space(3), " ", CStr(u)) )
        Set cmdArray(u).CmdDel = cmd_dict( replace(space(4), " ", CStr(u)) )

        Set PRange = ActiveCell.CurrentRegion
        Dim fields_available As Range
        Set fields_available = a_wkb.ActiveSheet.Range( _
            Cells(PRange.Rows(1).row, PRange.Columns(1).column).Address, _
            Cells(PRange.Rows(1).row, PRange.Columns(PRange.Columns.count).column).Address)

        if not fields_available.value = empty then levels1(u).list = transposeArray(fields_available.value)
        u = u + 1

    Loop While False: Next head

    form2.CommandButton1.Top = levels1(u-1).top + levels1(u-1).Height + 30   
    Update_ListBox

    form2.Width = 500
    form2.Height = 500
    form2.Show

End Function

Public Function MoveItem_Up()
    
    dim i, cell as range
    dim box_number as Integer: box_number = cint(right(form2.activecontrol.name,1))

    With wkb.Sheets(opt.Value)
        For i = 1 To levels2(box_number).ListCount - 1
            If levels2(box_number).Selected(i) Then
                'Edit personal sheet
                Set cell = .Range(.Cells(1, box_number).Address, .Cells(20, box_number).Address).Find( _
                    what:=levels2(box_number).list(i), _
                    LookAt:=xlWhole)
                cell.Offset(-1, 0).Insert shift:=xlDown
                cell.Cut (.Range(cell.Offset(-2, 0).Address))
                cell.Offset(2, 0).Delete shift:=xlUp
    
            End If
        Next i
        
        'Edit listbox
        Update_ListBox
        
    End With
    

end Function

Public Function MoveItem_Down()
    
    dim i, cell as range
    dim box_number as Integer: box_number = cint(right(form2.activecontrol.name,1))

    With wkb.Sheets(opt.Value)
        For i = levels2(box_number).ListCount - 2 To 0 step -1
            If levels2(box_number).Selected(i) Then
                'Edit personal sheet
                Set cell = .Range(.Cells(1, box_number).Address, .Cells(20, box_number).Address).Find( _
                    what:=levels2(box_number).list(i), _
                    LookAt:=xlWhole)
                cell.Offset(2, 0).Insert shift:=xlDown
                cell.Cut (.Range(cell.Offset(2, 0).Address))
                cell.Offset(-2, 0).Delete shift:=xlUp
    
            End If
        Next i
        
        'Edit listbox
        Update_ListBox
        
    End With
    
end Function

Public Function Add_Item()
    
    dim i
    dim box_number as Integer: box_number = cint(right(form2.activecontrol.name,1))

    For i = 0 To levels1(box_number).ListCount - 1
        If levels1(box_number).Selected(i) = True Then 
            levels2(box_number).AddItem levels1(box_number).List(i)
            
            with wkb.Worksheets(opt.value)
                .range(.cells( _
                    .cells(20,box_number).end(xlUp).row + 1, box_number).Address).value = levels1(box_number).List(i)
            end with
        
        end if
    Next i

End Function

Public Function Delete_Item()

    Dim counter As Integer: counter = 0
    dim i, cell as range
    dim box_number as Integer: box_number = cint(right(form2.activecontrol.name,1))
    
    For i = 0 To levels2(box_number).ListCount - 1
        If levels2(box_number).Selected(i - counter) Then
            
            with wkb.Worksheets(opt.value)
                set cell = .range(.cells(1,box_number).address, .cells(20,box_number).address).find( _
                    what:=levels2(box_number).List(i - counter), _
                    LookAt:=xlWhole)
                cell.delete(xlShiftUp)
            end with
            
            levels2(box_number).RemoveItem (i - counter)
            counter = counter + 1
        End If
    
    Next i

End Function

Public Function Update_ListBox()
    
    Set wkb = Application.Workbooks("PERSONAL.xlsb")
    
    
    Dim Rng As Range, Dn As Range
    Set Dic = CreateObject("scripting.dictionary")
    Dic.CompareMode = vbTextCompare

    Dim o As Integer
    For o = 1 To 4
        
        With wkb.Worksheets(opt.value)
            set Rng = .Range(.Cells(2, o).Address, .Cells(.Cells(20, o).End(xlUp).row, o).Address)
            For Each Dn In Rng
                If Not Dn.Value = "" Then
                    Set Dic(Dn.Value) = Dn
                End If
            Next
        
            If .Cells(20, o).End(xlUp).row > 2 Then
                levels2(o).list = .Range(.Cells(2, o).Address, .Cells(.Cells(20, o).End(xlUp).row, o).Address).value
            ElseIf .Cells(20, o).End(xlUp).row = 2 Then
                Dim arr(1) As String
                levels2(o).Clear
                levels2(o).AddItem .Range(.Cells(2, o).Address, .Cells(.Cells(20, o).End(xlUp).row, o).Address).value
            Else
                levels2(o).Clear
            End If
        End With
        
    Next o

End Function
