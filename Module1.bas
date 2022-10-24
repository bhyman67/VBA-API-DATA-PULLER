Option Explicit

' Important Notes!!!
'   -> JsonConvert.bas from https://github.com/VBA-tools/VBA-JSON needs
'      to be imported into this VBA project
'   -> Microsoft Scripting Runtime has to be added as a reference

' Declare variables and instantiate objects
Dim key, coll_itm, crnt_row, org_id, current_key
Dim hReq As Object
Dim JSON_Resp As Dictionary
Dim temp_dict As Dictionary
Dim strUrl As String
Dim response As String
Dim api_data_sht, ui_sht As Worksheet
Dim lst_row As Integer
Dim coll_itm_count As Integer

' Traverse through the JSON (recursivly) and write it to the spreadsheet
Private Sub traverse_and_write_json(JSON_Resp As Dictionary, prev_keys As String)

    ' account for when it's a list, dict, and key value pair
    For Each key In JSON_Resp

        ' Check for whether or not the itm is a dict
        If TypeName(JSON_Resp(key)) = "Dictionary" Then

            traverse_and_write_json JSON_Resp:=JSON_Resp(key), prev_keys:=prev_keys + key + "."

        ElseIf TypeName(JSON_Resp(key)) = "Collection" Then
            
            ' Loop through the collection
            coll_itm_count = 1
            current_key = key
            For Each coll_itm In JSON_Resp(current_key)
            
                If TypeName(coll_itm) = "Dictionary" Then
                
                    Set temp_dict = coll_itm
                    traverse_and_write_json JSON_Resp:=temp_dict, prev_keys:=prev_keys + current_key + "[" + CStr(coll_itm_count) + "]" + "."
                
                Else
                
                    With api_data_sht.Range("A" + CStr(crnt_row))
                        .Value = prev_keys + current_key + "[" + CStr(coll_itm_count) + "]"
                        .Offset(0, 1).Value = coll_itm
                    End With
                    crnt_row = crnt_row + 1
                
                End If
                coll_itm_count = coll_itm_count + 1
            
            Next coll_itm

        Else

            With api_data_sht.Range("A" + CStr(crnt_row))
                .Value = prev_keys + key
                .Offset(0, 1).Value = JSON_Resp(key)
            End With
            crnt_row = crnt_row + 1

        End If

    Next key

End Sub

Sub Main()

    ' Set-up worksheet objects
    Set api_data_sht = Worksheets("API Data")
    Set ui_sht = Worksheets("UI")
    
    ' Delete any previous data
    api_data_sht.Range("A:B").Delete shift:=xlShiftToLeft

    ' API endpoint
    strUrl = Trim(ui_sht.Range("E6").Value)

    ' Web request to API endpoint
    ' -> Is there a standard way to make web requests in vba???
    Set hReq = CreateObject("MSXML2.XMLHTTP")
    With hReq
        .Open "GET", strUrl, False
        .Send
    End With

    ' Grab JSON resp from API endpoint and parse it
    response = hReq.ResponseText
    Set JSON_Resp = JsonConverter.ParseJson(response)

    ' Run the traverse and write json routine
    crnt_row = 1 ' This will be a global var...
    traverse_and_write_json JSON_Resp:=JSON_Resp, prev_keys:=""
    
    ' Auto-fit and align values to the left
    api_data_sht.Columns("A:B").AutoFit
    api_data_sht.Range("B:B").HorizontalAlignment = xlLeft
    
    ' Done...
    MsgBox ("Done")

End Sub
