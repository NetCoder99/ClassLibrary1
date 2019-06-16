' --------------------------------------------------------------
' Testing for calling .Net Framework from vbs
' --------------------------------------------------------------
Call CheckArguments
Call GetHelloTest1
Call GetHelloTest2

' --------------------------------------------------------------
' Sub  : GetHelloTest2
' Desc : Create the COM object and call a function that 
'        returns an array of strings
' --------------------------------------------------------------
Sub GetHelloTest2
    Dim objObject
    Set objObject = WScript.CreateObject("ClassLibrary1.Class1")
    Dim objOutput
    Set objOutput = objObject.HelloList

    For each item in objOutput
        WScript.Echo item
    Next
    '?? WScript.Echo objOutput ??
End Sub


' --------------------------------------------------------------
' Sub  : GetHelloTest1
' Desc : Create the COM object and call a simple function
' --------------------------------------------------------------
Sub GetHelloTest1
    Dim objObject
    Set objObject = WScript.CreateObject("ClassLibrary1.Class1")
    Dim objOutput
    objOutput = objObject.Hello
    WScript.Echo objOutput
End Sub


' --------------------------------------------------------------
' Sub  : CheckArguments
' Desc : Verify the arguments and create a dictionary object
' --------------------------------------------------------------
Sub CheckArguments
    text = "Arguments" & vbCrLf & vbCrLf
    Set objArgs = WScript.Arguments       ' Create object.
    For i = 0 to objArgs.Count - 1        ' Loop through all arguments.
        text = text & objArgs(i) & vbCrLf ' Get argument.
    Next 
    WScript.Echo text ' Show arguments using Echo.
End Sub
