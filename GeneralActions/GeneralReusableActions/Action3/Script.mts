Dim browser
Dim url

' *** This section is working with Action parameters
' browser = Parameter.Item("p_BrowserType")
' url = Parameter.Item("p_URL")

' ALM Test Configurations seem to work more with DataTable, so I use DataTable here
browser = DataTable("Browser", dtGlobalSheet)
url = DataTable("URL", dtGlobalSheet)

If  browser = "IE" Then
	browser = BROWSER_IE
ElseIf browser = "CHROME" Then
	browser = BROWSER_CHROME
ElseIf browser = "FIREFOX" Then
	browser = BROWSER_FIREFOX
Else 
	browser = "DEFAULT" 
End If

LaunchBrowserNavigateToURL browser, url

