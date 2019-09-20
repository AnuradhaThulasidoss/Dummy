Dim tdc


Dim Name, Password

Name = ""
Password = ""

Set tdc = CreateObject("TDApiOle80.TDConnection")
tdc.InitConnectionEx "http://sjalmappprdn03:8080/qcbin"
tdc.Login Name, Password
tdc.Connect "MOBILE_WEB_AND_PORTALS", "GForce"

If (tdc.connected = True) Then
MsgBox "Connected to ALM Project"
'WScript.Quit
End If

Set tfact = tdc.TestSetFactory
Set tsTreeMgr = tdc.TestSetTreeManager
Set tcTreeMgr = tdc.TreeManager
nPath = "Root\" & Trim("2018\Dry Run")
Set TestSetFolder = tsTreeMgr.NodeByPath(npath)

Set TestSetF = TestSetFolder.TestSetFactory 'Retreive test from given folder in test lab
Set aTestSetArray = TestSetF.NewList("")
tsSet_cnt=aTestSetArray.Count
For i=1 to tsSet_cnt ' Loop through the Test Sets to pick the desired test Set
	Set tstests=aTestSetArray.Item(i)
	TestSet_Name=tstests.Name
	'MsgBox TestSet_Name
	If TestSet_Name = "LabNoteBook_Lightning" Then 
        Set Scheduler = tstests.StartExecution("")
        Scheduler.RunAllLocally = True
        Scheduler.Run(Test)
        Set execStatus = Scheduler.ExecutionStatus
        Exit For
    End If
Next 