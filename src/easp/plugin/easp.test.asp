<%
'#################################################################################
'##	easp.test.asp
'##	------------------------------------------------------------------------------
'##	Feature		:	EasyAsp Plugin Class Sample
'##	Version		:	v0.1
'##	Author		:	Coldstone(coldstone[at]qq.com)
'##	Update Date	:	2009/12/3 12:23
'##	Description	:	EasyAsp's plugin should be like this file as follow:
'					1.	File name should be like this: 'easp.***.asp'.  The '***'
'						is your plugin's name, with lower-case letters as better.
'					2.	Class's name should be like this: 'EasyAsp_***'. The '***'
'						is your plugin's name, lower-case letters after the '_'
'						are not required.
'					3.	You must put your file(s) in 'plugin' folder or any other
'						folder you setted with the property 'Easp.PluginPath'.
'#################################################################################
Class EasyAsp_Test

	Private s_author, s_version

	Private Sub Class_Initialize()
		s_author	= "coldstone"
		s_version	= "0.1"
		'Set Exception Info, please custom your ErrorCode such as 10001 or "test001".
		'Please try to keep your ErrorCode is not the same with other people.
		'Attention! You can use any character, not just numbers!
		Easp.Error(10001) = "Error! Accept number only."
	End Sub
	Private Sub Class_Terminate()
		
	End Sub
	'Set Property
	Public Property Get Author()
		Author = s_author
	End Property
	Public Property Get Version()
		Version = s_version
	End Property
	'Define a 'Sub'
	Public Sub helloWorld()
		Easp.W "Hello World!"
	End Sub
	'Define a 'Function'
	Public Function Return(ByVal s)
		Return = s
	End Function
	'Define default Function
	Public Default Function Fun(ByVal num)
		If Not isNumeric(num) Then
			'Raise a Exception Info When 'Easp.Debug' is 'True'
			Easp.Error.Msg = "<br />Your input is: " & num
			Easp.Error.Raise 10001
		End If
		Fun = num
	End Function

End Class
%>