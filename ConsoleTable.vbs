Option Explicit

Class ConsoleTable  'Defining the Class

    Private objDic
    Private strHeads
    Private strWidths
    Private strSpacing
    Private intIndexer
    
    Private Sub Class_Initialize(  )
        Set objDic = WScript.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate(  )
        Set objDic = Nothing
    End Sub

    Private Function GetTableWidth()    
        GetTableWidth = Len(strWidths) + 1
    End Function

    Sub SetHeaders(strHeaders)
        strHeads = strHeaders
        UpdateWidths(strHeaders)
    End Sub

    Private Function FormatInput(strInputRow)   
        Dim InputRowArray
            InputRowArray = Split(strInputRow,",")
        Dim intIndexer, strTmpOutput
        For intIndexer = 0 to UBound(InputRowArray)
            strTmpOutput = strTmpOutput & Trim(InputRowArray(intIndexer)) & "," 
        Next
        FormatInput = strTmpOutput
    End Function

    Private Sub UpdateWidths(strNewRowContent)      
        Dim NewRowArray : NewRowArray = Split(strNewRowContent,",")
        If UBound(NewRowArray) > 0 Then
            SetNewWidths NewRowArray
        End If
    End Sub

    Private Sub GetNewWidths(NewArray)
        Dim strTemp
        For intIndexer = 0 to UBound(NewArray)
            intWidth = Len(strPreviousWidth)
            intNewWidth = Len(strNewWidth)
            If intNewWidth > intWidth Then
                intWidth = intNewWidth
            End If
            Select Case intIndexer
                Case 0
                    strWidths = String(intWidth,"X") & ","
                Case UBound(WidthsArray)
                    strWidths = strWidths & String(intWidth,"X")
                Case Else
                    strWidths = strWidths & String(intWidth,"X") & ","
            End Select
        Next
    End Function

    Sub AddRow(strRowContent)
        Dim strFormated : strFormated = FormatInput(strRowContent)
        objDic.Add objDic.Count, strFormated
        UpdateWidths(strFormated)
    End Sub

    Private Sub WriteRow(strRowContent)
        Dim strTemp
        Dim intSpaces
        Dim intLSpaces
        Dim intRSpaces
        Dim strCellContent
        ' Retrieve columns' widths
        Dim WidthArray
            WidthArray = Split(strWidths, ",")
        ' Parse content of each cell
        Dim ContentArray
            ContentArray = Split(strRowContent, ",")
        Dim SpacingArray
            SpacingArray = Split(strSpacing, ",")
        If UBound(WidthArray) > 0 Then
            'Loop through the width's array
            For intIndexer = 0 to UBound(WidthArray)
                ' Get the cell content
                If intIndexer > UBound(ContentArray) Then
                    strCellContent = ""
                Else
                    strCellContent = ContentArray(intIndexer)
                End If
                ' calculate spacing inside each cell
                intSpaces = Len(WidthArray(intIndexer)) - Len(strCellContent)
                Select Case intSpacingMode
                    Case 0 ' Left
                        strTemp = strTemp & strCellContent
                        strTemp = strTemp & Space(intSpaces) & GetVerticalBorder
                    Case 2 ' Rigth
                        strTemp = strTemp & Space(intSpaces)
                        strTemp = strTemp & strCellContent & GetVerticalBorder
                    Case Else ' Center
                        intLSpaces = Int(intSpaces/2)
                        intRSpaces = intSpaces - intLSpaces
                        strTemp = strTemp & Space(intLSpaces) & strCellContent
                        strTemp = strTemp & Space(intRSpaces) & GetVerticalBorder
                End Select
            Next
        End If
        ' End of Row
        WScript.Echo(strTemp)   ' Write Row
    End Sub

    Sub Write 
        Dim Items
            Items = objDic.Items
        
        'Print start of table
        WScript.Echo(String(GetTableWidth,HorizontalMark))

        If Len(strHeads) > 0 Then                
            'Print headers
            WriteRow(strHeads)

            'Print separator
            WScript.Echo(String(GetTableWidth,HorizontalMark))
        End If

        If UBound(Items) > 0 Then
            'Print rows
            Dim intRow
            For intRow = 0 to UBound(Items)
                WriteRow(Items(intRow))        
            Next
        End If

        'Print end of table
        WScript.Echo(String(GetTableWidth,HorizontalMark))
    End Sub

    

End Class

' Instantiation of the Class
Dim objTable
Set objTable = New ConsoleTable

With objTable
    .SetHeaders("Numero,Pais,Fondos")
    .AddRow("1,Cuba,1000")
    .AddRow("2,Rusia,10000")
    .AddRow("3,Estados Unidos,1000000")
    .AddRow("4,Brasil")
    .AddRow("5,China, 9999999999999")
    .AddRow("6, EU , 7777777, qwerty")
    .AddRow("7, Portugal , 333x7777777, qwerty")
    .Write
End With
