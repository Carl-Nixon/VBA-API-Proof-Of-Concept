Attribute VB_Name = "API_POC"
' This is a proof of concept tool for interacting with APIs via VBA
'
' The programming calls upon the following;
' 1. The Json.Converter.bas module from the VBA-JSON project on GitHHub. This can be found at;
'    https://github.com/VBA-tools/VBA-JSON
' 2. A reference to the "Microsoft WinHTTP Services, version 5.1" library
' 3. A reference to the "Microsoft Scripting Runtime" library
'
Sub Financial_Data()

    ' This calls upon the Alpha Vantage API to pull share price information
    ' Keys and API documentation for this API can be accessed here;
    ' https://rapidapi.com/alphavantage/api/alpha-vantage/
    
    ' +---------------------------------------------------------+
    ' | This function returns the opening and closing positions |
    ' | of a given share                                        |
    ' +---------------------------------------------------------+
    
    ' >>>> Set Up Variables <<<<
    Dim URL As String                       ' Holds query URL
    Dim Parameter_interval As String        ' Holds the parameter for the time intervals
    Dim Parameter_function As String        ' Holds the query function being called on
    Dim Parameter_symbol As String          ' Holds the stock being searched for
    Dim Parameter_datatype As String        ' Holds the format for the output
    Dim Parameter_output_size As String     ' Holds the size of the output we want
    Dim Request As New WinHttpRequest       ' Holds our request to the API
    Dim API_Host As String                  ' This API needs a host
    Dim API_Key As String                   ' Holds the API key
    Dim Response As Object                  ' Holds the response from the API
    Dim Meta_Data As Dictionary             ' Holds the Meta Data from response
    Dim Time_Series As Dictionary           ' Holds the Time Series data in a dictionary
    
    
    URL = "https://alpha-vantage.p.rapidapi.com/query"  ' Set query URL
    
    Parameter_interval = "5min"                 ' Can be 1, 5, 15, 30 or 60 minutes
    Parameter_function = "TIME_SERIES_INTRADAY" ' Name of the query function being called upon
    Parameter_symbol = "MSFT"                   ' Short code for the stock being checked
    Parameter_datatype = "json"                 ' Type of dataset we want to return. json / csv
    Parameter_output_size = "compact"           ' Compact returns the last 100 records.
                                                ' Full returns the full set
                                                
    ' Get API Host & Key Get from https://rapidapi.com/alphavantage/api/alpha-vantage/
    API_Host = "Get from https://rapidapi.com/"
    API_Key = "Get from https://rapidapi.com/"

    ' >>>> Prepare the results tab <<<<
    Sheets("Results").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("Results").Cells(1, 1) = "1. Information"
    Sheets("Results").Cells(2, 1) = "2. Symbol"
    Sheets("Results").Cells(3, 1) = "3. Last Refreshed"
    Sheets("Results").Cells(4, 1) = "4. Interval"
    Sheets("Results").Cells(5, 1) = "5. Output Size"
    Sheets("Results").Cells(6, 1) = "6. Time Zone"
    Sheets("Results").Cells(7, 1) = "Date / Time"
    Sheets("Results").Cells(7, 2) = "Open"
    Sheets("Results").Cells(7, 3) = "High"
    Sheets("Results").Cells(7, 4) = "Low"
    Sheets("Results").Cells(7, 5) = "Close"
    Sheets("Results").Cells(7, 6) = "Volume"
    
    ' >>>> Prepare API query <<<<
    
    ' Add interval to URL
    URL = URL & "?interval=" & CStr(Parameter_interval)
    
    ' Add function to URL
    URL = URL & "&function=" & Parameter_function
    
    ' Add symbol to URL
    URL = URL & "&symbol=" & Parameter_symbol
    
    ' Add datatype to URL
    URL = URL & "&datatype=" & Parameter_datatype
    
    ' Add output size to URL
    URL = URL & "&output_size=" & Parameter_output_size
    
    ' >>>> Send the Request <<<<
    
    ' Open the request
    Request.Open "GET", URL
    
    ' Add host to request header
    Request.SetRequestHeader "x-rapidapi-host", API_Host
    
    ' Add api key to header
    Request.SetRequestHeader "x-rapidapi-Key", API_Key
    
    ' Send the request
    Request.Send
    
    ' >>>> Make sure there was a response <<<<
    
    ' If a non 200 response was received there was a problem
    If Request.Status <> 200 Then
        
        ' Display details of the response / error
        MsgBox Request.ResponseText
        
        Exit Sub
    End If
    
    ' >>>> Process Response (Dictionary) <<<<
    ' Capture the response
    Set Response = JsonConverter.ParseJson(Request.ResponseText)
    ' Put the Meta data in to a dictionary
    Set Meta_Data = Response("Meta Data")
    ' Put meta data into results tab
    Sheets("Results").Cells(1, 2) = Meta_Data("1. Information")
    Sheets("Results").Cells(2, 2) = Meta_Data("2. Symbol")
    Sheets("Results").Cells(3, 2) = Meta_Data("3. Last Refreshed")
    Sheets("Results").Cells(4, 2) = Meta_Data("4. Interval")
    Sheets("Results").Cells(5, 2) = Meta_Data("5. Output Size")
    Sheets("Results").Cells(6, 2) = Meta_Data("6. Time Zone")
    
    ' Set the row to insert the data into
    InsertRow = 8
    
    ' >>>> Process Time Series (Dictionary of Dictionaries) <<<<
    
    ' Capture the time series data in a dictionary
    Set Time_Series = Response("Time Series (" & Parameter_interval & ")")
    
    ' Loop through Time_Series data (outer dictionary)
    For Each oKey In Time_Series.Keys
        
        ' Put the time stamp in the results
        Sheets("Results").Cells(InsertRow, 1) = oKey
        
        ' Set column to start putting data into
        InsertCol = 2
        
        ' Loop through the inner dictionary
        For Each iKey In Time_Series(oKey)
            
            ' Put data into results tab
            Sheets("Results").Cells(InsertRow, InsertCol) = Time_Series(oKey)(iKey)
            
            ' Increment insert column
            InsertCol = InsertCol + 1
        
        Next iKey
        
        ' Increment insert row
        InsertRow = InsertRow + 1
    
    Next oKey
    
    ' >>>> Display Confirmation Message <<<<
    MsgBox "Finished"

End Sub
