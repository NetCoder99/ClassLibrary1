' --------------------------------------------------------------
' Testing for calling .Net Framework from vbs
' --------------------------------------------------------------
Call CheckArguments
Call GetHelloTest1
Call GetHelloTest2
Call GetPerson
Call GetPersonList
Call GetPersonListJSON

' --------------------------------------------------------------
' Sub  : GetPersonListJSON
' Desc : Create the COM object and call a function that 
'        returns a JSON string. Test if dll will find 
'        NewtonSoft library.
' --------------------------------------------------------------
Sub GetPersonListJSON
    WScript.Echo "-----" + "GetPersonListJSON" + "-----"
    Dim objObject
    Set objObject = WScript.CreateObject("ClassLibrary1.Class1")
    Dim objOutput
    objOutput = objObject.GetPersonListJSON
    WScript.Echo  objOutput
End Sub


' --------------------------------------------------------------
' Sub  : GetPerson
' Desc : Create the COM object and call a function that 
'        returns a C# class.
' --------------------------------------------------------------
Sub GetPersonList
    WScript.Echo "-----" + "GetPersonList" + "-----"
    Dim objObject
    Set objObject = WScript.CreateObject("ClassLibrary1.Class1")
    Dim objOutput
    Set objOutput = objObject.GetPersonList
    For each item in objOutput
        WScript.Echo  CStr(item.Id) + ";" + item.FName + ";" + item.LName 
    Next
End Sub

' --------------------------------------------------------------
' Sub  : GetPerson
' Desc : Create the COM object and call a function that 
'        returns a C# class.
' --------------------------------------------------------------
Sub GetPerson
    Dim objObject
    Set objObject = WScript.CreateObject("ClassLibrary1.Class1")
    Dim objOutput
    Set objOutput = objObject.GetPerson
    WScript.Echo  CStr(objOutput.Id) + ";" + objOutput.FName + ";" + objOutput.LName
End Sub

' --------------------------------------------------------------
' Sub  : GetHelloTest2
' Desc : Create the COM object and call a function that 
'        returns an array of strings
' --------------------------------------------------------------
Sub GetHelloTest2
    WScript.Echo "-----" + "GetHelloTest2" + "-----"
    Dim objObject
    Set objObject = WScript.CreateObject("ClassLibrary1.Class1")
    Dim objOutput
    Set objOutput = objObject.HelloList
    For each item in objOutput
        WScript.Echo item
    Next
End Sub


' --------------------------------------------------------------
' Sub  : GetHelloTest1
' Desc : Create the COM object and call a simple function
' --------------------------------------------------------------
Sub GetHelloTest1
    WScript.Echo "-----" + "GetHelloTest1" + "-----"
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
    WScript.Echo "-----" + "CheckArguments" + "-----"
    Set objArgs = WScript.Arguments       ' Create object.
    For i = 0 to objArgs.Count - 1        ' Loop through all arguments.
        WScript.Echo objArgs(i)
    Next 
    'WScript.Echo text ' Show arguments using Echo.
End Sub
