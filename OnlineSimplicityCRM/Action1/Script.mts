Browser("Simplicity").Page("Simplicity").WebEdit("user_name").Set "admin"
Browser("Simplicity").Page("Simplicity").WebEdit("user_password").SetSecure "56e765ea2687e971bdcebd57cbee07a5025bdd557b62f6d3d7a0f9b5"
Browser("Simplicity").Page("Simplicity").WebButton("Log In").Click
Browser("Simplicity").Page("Simplicity").Link("Log Out").Click
