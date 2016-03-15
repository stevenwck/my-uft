Dim userName 
Dim password

' userName = Parameter.Item("p_UserName")
' password = Parameter.Item("p_Password")

userName = DataTable("UserName", dtGlobalSheet)
password = DataTable("Password", dtGlobalSheet)

' Check that the Logo shown on screen is as expected
Browser("Simplicity").Page("Simplicity").Image("Sugar").Check CheckPoint("Sugar")

Browser("Simplicity").Page("Simplicity").WebEdit("user_name").Set userName
Browser("Simplicity").Page("Simplicity").WebEdit("user_password").SetSecure password
Browser("Simplicity").Page("Simplicity").WebButton("Log In").Click

' Check that we are looking at the Home Page.
Browser("Simplicity").Page("Simplicity").Link("Home").Check CheckPoint("Home") @@ hightlight id_;_Browser("Simplicity").Page("Simplicity 2").Link("Home")_;_script infofile_;_ZIP::ssf3.xml_;_

 @@ hightlight id_;_Browser("Simplicity").Page("Simplicity").Image("Sugar")_;_script infofile_;_ZIP::ssf4.xml_;_
