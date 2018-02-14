Attribute VB_Name = "S9_References"
Function getDictPatentInfobyPatNum(ByVal patnum As String) As Dictionary
    Dim html As New HTMLDocument
    Dim url As String
    Dim XMLHTTP As Object
    Dim start_time As Date
    Dim end_time As Date
    start_time = Time
    Debug.Print "===start_time:" & start_time & "==="
    url = "https://www.google.com/patents/" & patnum & "?hl=en-us"

    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
    XMLHTTP.Open "GET", url, False
    XMLHTTP.setRequestHeader "Content-Type", "text/xml"
    XMLHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; rv:25.0) Gecko/20100101 Firefox/25.0"
    XMLHTTP.send

    Do While Not XMLHTTP.readyState = 4
        DoEvents
    Loop

    html.Body.innerHTML = XMLHTTP.responseText

    Set bibdata = html.getElementsByClassName("patent-bibdata")(0) 'object
    Set topic = bibdata.getElementsByTagName("tr") ''colection

    Dim PatentDocInfo As New Dictionary
    For i = 0 To 8
        strKey = Trim(topic(i).getElementsByTagName("td")(0).innerText)
        infoValue = Trim(topic(i).getElementsByTagName("td")(1).innerText)
        If Not PatentDocInfo.Exists(strKey) Then PatentDocInfo.Add strKey, infoValue
        Debug.Print strKey, ": ", Tab, PatentDocInfo(strKey)

    Next
    
    Set getDictPatentInfobyPatNum = PatentDocInfo
    end_time = Time
    Debug.Print "===end_time:" & end_time & "==="
    Debug.Print "Done, Time taken : " & DateDiff("n", start_time, end_time)
End Function
