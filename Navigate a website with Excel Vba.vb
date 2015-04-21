Sub googlesearch()
    Set objIE = CreateObject("InternetExplorer.Application")
    WebSite = "www.google.com"
    With objIE
        .Visible = True
        .navigate WebSite
        Do While .Busy Or .readyState <> 4
            DoEvents
        Loop

        Set Element = .document.getElementsByName("q")
        Element.Item(0).Value = "Hello world"
        .document.forms(0).submit
        '.quit
        End With

End Sub

''http://stackoverflow.com/questions/2632306/navigate-a-website-with-excell-vba
