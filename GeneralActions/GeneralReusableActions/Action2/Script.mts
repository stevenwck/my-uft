Browser("Simplicity").Page("Simplicity").Link("Log Out").Click
'Browser("Simplicity").Page("Simplicity").Image("Sugar").Check CheckPoint("Sugar")

' Check for Log In button to be on the screen
Browser("Simplicity").Page("Simplicity").WebButton("Log In").Check CheckPoint("Log In")

Browser("Simplicity").Close
