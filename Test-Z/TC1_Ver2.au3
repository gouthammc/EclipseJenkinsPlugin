#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.10.2
 Author: Rana Banerjee

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

#include <Excel.au3>
#include <Date.au3>
#include <IE.au3>
#include <ReusableFunctions.au3>

; Script Start ***************************************************************************************************************************************
;Initialization
;Initialization function
Func _initialize()
Global $iteration=$CmdLine[1]
Global $hTimer = TimerInit()
Global $TestCaseName = _GetData($iteration, "TestCaseName")
Global $TestCaseDescription = _GetData($iteration, "TestCaseDescription")
;initialize the result file
Global $oResExcel = _ExcelBookNew (0)
_ExcelBookSaveAs($oResExcel,@ScriptDir & "\Results\Result_"& $TestCaseName & "_" & StringReplace(_DateTimeFormat(_NowCalc(), 2),"-", "_") & "_"& StringReplace(_DateTimeFormat(_NowCalc(), 4),":", "_"))
_ExcelSheetAddNew ($oResExcel,"Results")
_ExcelWriteCell($oResExcel, "Test Case Name:", 1, 1)
_ExcelWriteCell($oResExcel, $TestCaseName, 1, 2)
_ExcelWriteCell($oResExcel, "Test Case Description:", 2, 1)
_ExcelWriteCell($oResExcel, $TestCaseDescription, 2, 2)
_ExcelWriteCell($oResExcel, "Iteration:", 3, 1)
_ExcelWriteCell($oResExcel, $iteration, 3, 2)
Global $ResultRow = 4
_ResultReporter ("Initialization Complete",0)
EndFunc

_initialize()

_ResultReporter ("Run the eclipse executable",0)
Run (_GetData($iteration, "EclipseExePath"))
_ResultReporter ("Wait for the Workspace launcher dialog box",0)
_WinWaitActiveMod ("Workspace Launcher","Eclipse stores your projects",240)
_ResultReporter ("Enter Workspace name in the Workspace Launcer",0)
ControlSetText ("Workspace Launcher","Eclipse stores your projects", "[CLASS:Edit;INSTANCE:1]", _GetData($iteration, "JavaWorkspacePath"))
Sleep (1000)
_ResultReporter ("Click on the OK button",0)
ControlClick ("Workspace Launcher","Eclipse stores your projects","[CLASS:Button;INSTANCE:3]")
_ResultReporter ("Wait for the Eclipse Window to open",0)
_WinWaitActiveMod ("Java", "", 240)

_ResultReporter ("Maximize the window",0)
Local $hWnd = WinWait("Java", "", 10)
WinSetState($hWnd, "", @SW_MAXIMIZE)

_WinWaitActiveMod ("Java", "", 240)
sleep(2000)
_ResultReporter ("Send keystokes to Open Perspective",0)
Send ("!woo")

_WinWaitActiveMod ("Open Perspective", "", 120)
_ResultReporter ("Send Keystrokes to Open the Java - EE perspective",0)
Send ("jjj")
ControlClick ("Open Perspective", "","[CLASS:Button;Text:OK]")
_WinWaitActiveMod ("Java EE - ","",20)
_ResultReporter ("Click on the button to close the introductory page if present",0)
ControlClick ("Java EE - ", "", "[CLASS:SWT_Window0;INSTANCE:3]","left",1,90,14)
_WinWaitActiveMod ("Java EE - ", "", 240)
Sleep (2000)
_ResultReporter ("Send Keystrokes to open the dialog box for creation of standalone dynamic web project",0)
Send ("!fnd")
_WinWaitActiveMod ("New Dynamic Web Project", "Create a standalone Dynamic Web project", 15)
_ResultReporter ("Give a name to the project",0)
ControlSetText ("New Dynamic Web Project", "Create a standalone Dynamic Web project", "[CLASS:Edit;INSTANCE:1]",_GetData($iteration, "DynamicWebProjectName"))
_ResultReporter ("Click on the Finish button",0)
ControlClick ("New Dynamic Web Project", "Create a standalone Dynamic Web project", "[CLASS:Button;Text:&Finish]")
_ResultReporter ("Give sufficient time for the project to be created: 10 seconds in this Case",0)
Sleep (10000)
_WinWaitActiveMod ("Java EE - ", "", 120)
_ResultReporter ("Select the newly created project in the project explorer",0)
ControlClick ("Java EE -","","[CLASS:SysTreeView32;INSTANCE:1]","left",1,50,10)
sleep (2000)
_ResultReporter ("Send Keystorkes to go to the Web folder, right click (via Shift+F10) and choose create new JSP file option",0)
_WinWaitActiveMod ("Java EE -", "", 10)
Send ("{RIGHT}")
Send ("{DOWN 6}")
Send ("{SHIFTDOWN}")
Send ("{F10}")
Send ("{SHIFTUP}")
Send ("{DOWN}")
Send ("{RIGHT}")
Send ("{DOWN 5}")
Send ("{ENTER}")
_WinWaitActiveMod ("New JSP File", "Create a new JSP file", 60)
ControlSetText ("New JSP File", "Create a new JSP file","[CLASS:Edit;INSTANCE:2]",_GetData($iteration, "JSPName"))
ControlClick ("New JSP File", "Create a new JSP file","[CLASS:Button;Text:&Next >]")
ControlClick ("New JSP File", "Select JSP Template","[CLASS:Button;Text:&Finish]")
sleep (5000)

_ResultReporter ("go to the 9th line and enter Hello world text",0)
_WinWaitActiveMod ("Java EE - ", "", 120)
Send ("{DOWN 9}")
Send (_GetData($iteration,  "JSPText"))
Send("{SHIFTDOWN}")
Send("{END}")
Send ("{SHIFTUP}")
Send ("{DEL}")
_ResultReporter ("Save index file",0)
Send ("{CTRLDOWN}")
Send ("s")
Send ("{CTRLUP}")
_ResultReporter ("give some time for the file to be saved",0)
Sleep (5000)
_WinWaitActiveMod ("Java EE - ", "", 120)
_ResultReporter ("Highlight the project in the project explorer",0)
ControlClick ("Java EE -","","[CLASS:SysTreeView32;INSTANCE:1]","left",1,50,10)
_WinWaitActiveMod ("Java EE -", "", 10)
_ResultReporter ("Right click on the project and select the option to package it for Windows Azure project",0)
_WinWaitActiveMod ("Java EE -", "", 10)
ControlClick ("Java EE -","","[CLASS:SysTreeView32;INSTANCE:1]","right",1,50,10)
sleep(1000)
Send("e")
sleep(500)
Send("{LEFT}")
sleep(500)
Send("{UP}")
sleep(500)
Send("{RIGHT}")
sleep(500)
Send("{ENTER}")
_ResultReporter ("Wait till the New Windows azure deployment project window appears",0)
_WinWaitActiveMod("New Azure Deployment Project","Enter project name",120)
ControlSetText ("New Azure Deployment Project","Enter project name","[CLASS:Edit;INSTANCE:1]",_GetData($iteration,  "AzureProjName"))
ControlClick ("New Azure Deployment Project","Enter project name","[CLASS:Button;Text:&Next >]")


_ResultReporter ("Check the jdk checkbox",0)
   ControlCommand ("New Azure Deployment Project","","[CLASS:Button;Text:Use the JDK from this file path for testing locally:]",_GetData($iteration, "CheckJDKOption"), "")
_ResultReporter ("jdk path",0)
   ControlSetText("New Azure Deployment Project","","[CLASS:Edit;INSTANCE:3]",_GetData($iteration, "JDKPath"))
   ControlClick("New Azure Deployment Project", "Specify the JDK to use for this deployment","[CLASS:Button;Text:Deploy my local JDK (auto-upload to cloud storage)]")

_ResultReporter ("Move to the next tab: Server",0)
ControlCommand ("New Azure Deployment Project","","[CLASS:SysTabControl32;INSTANCE:1]","TabRight", "")
_ResultReporter ("Check the server path checkbox",0)
ControlCommand ("New Azure Deployment Project","","[CLASS:Button;Text:Use the server from this file path for testing locally:]",_GetData($iteration,  "CheckLocalServer"), "")

ControlSetText ("New Azure Deployment Project","","[CLASS:Edit;INSTANCE:6]",_GetData($iteration, "ServerPath"))


_WinWaitActiveMod("New Azure Deployment Project","Select the type of server to include",60)
_ResultReporter ("Select the server:1 corresponds to Apache Tomcat 7",0)
ControlCommand ("New Azure Deployment Project","","[CLASS:ComboBox;INSTANCE:4]","SetCurrentSelection",Number(_GetData($iteration, "ServerNo")))

_ResultReporter ("Move to the next tab: Applications",0)
ControlCommand ("New Azure Deployment Project","","[CLASS:SysTabControl32;INSTANCE:1]","TabRight", "")

_ResultReporter ("Click on the Next button to go to select any additional optional features, if any",0)
_WinWaitActiveMod("New Azure Deployment Project","Specify the applications to use for this deployment",60)
ControlClick ("New Azure Deployment Project","Specify the applications to use for this deployment", "[CLASS:Button;Text:&Next >]")
sleep(2000)
_ResultReporter ("Click on the finish button",0)
_WinWaitActiveMod("New Azure Deployment Project","Select any additional optional features to enable for the default role in your project",60)
ControlClick ("New Azure Deployment Project","Select any additional optional features to enable for the default role in your project", "[CLASS:Button;Text:&Finish]")
_ResultReporter ("Give some time for the Azure project to be created. 10 seconds in this case",0)
Sleep (10000)


_ResultReporter ("Click on the deploy to emulator button",0)
$instance = 15
While $instance<20
_WinWaitActiveMod("Java EE","",60)
ControlClick ("Java EE","","[CLASS:SysTreeView32;INSTANCE:1]","left",1,50,10)
Sleep (1000)
ControlClick ("Java EE","","[CLASS:ToolbarWindow32;INSTANCE:"&$instance&"]","left",1,10,10)
Sleep(1000)
Send("{ESC}")
sleep(2000)
if StringInStr (ControlGetText("Java EE","", "[CLASS:Static;INSTANCE:1]"),"WindowsAzureProjectBuilder")>0 or StringInStr (ControlGetText("Java EE","", "[CLASS:Static;INSTANCE:2]"),"WindowsAzureProjectBuilder")>0 or StringInStr (ControlGetText("Java EE","", "[CLASS:Static;INSTANCE:3]"),"WindowsAzureProjectBuilder")>0 or StringInStr (ControlGetText("Java EE","", "[CLASS:Static;INSTANCE:4]"),"WindowsAzureProjectBuilder")>0 or StringInStr (ControlGetText("Java EE","", "[CLASS:Static;INSTANCE:5]"),"WindowsAzureProjectBuilder")>0 Then
   ExitLoop
endif
$instance=$instance+1
Wend

_ResultReporter ("Give some time for the console to start displaying the data, 10 seconds in this case",0)
sleep(10000)

_ResultReporter ("Open the find text window",0)
$instance=40
Do
WinActivate("Java EE","")
ControlClick ("Java EE - ","","[CLASS:SWT_Window0;INSTANCE:"&$instance&"]","right",1, 10, 10)
sleep(1000)
Send ("f")
ControlClick ("Java EE - ","","[CLASS:SWT_Window0;INSTANCE:"&$instance&"]","right",1, 10, 10)
sleep(1000)
Send ("f")
ControlClick ("Java EE - ","","[CLASS:SWT_Window0;INSTANCE:"&$instance&"]","right",1, 10, 10)
sleep(1000)
Send ("f")
$instance=$instance+1
until WinExists("Find/Replace","")

_WinWaitActiveMod ("Find/Replace","",10)
ControlSetText ("Find/Replace","","[CLASS:Edit;INSTANCE:1]","BUILD SUCCESSFUL")
ControlClick ("Find/Replace", "", "[CLASS:Button;Text:Fi&nd]")

_ResultReporter ("Wait until build sucessful message is not found or a timeout of 10 mins",0)
$timer = 0
do
   sleep (5000)
   $timer = $timer+5
   ControlClick ("Find/Replace", "", "[CLASS:Button;Text:Fi&nd]")
until StringCompare (ControlGetText ("Find/Replace", "", "[CLASS:Static;INSTANCE:3]"), "String Not Found") <> 0 or $timer=600

Sleep (2000)
_ResultReporter ("Close the Find window",0)
ControlClick ("Find/Replace", "", "[CLASS:Button;Text:Close]")
_ResultReporter ("give some time for the emulator to start, 10 seconds in this case",0)
Sleep (10000)

_ResultReporter ("Launch IE and open the localhost url",0)

Global $oIE
$timer=0
Do
   if WinExists("Internet Explorer","&Close program")=1 Then
	  ControlClick ("Internet Explorer","&Close program", "[CLASS:Button;Text:&Close program]")
	  ControlClick ("Internet Explorer","&Close program", "[CLASS:Button;Text:&Close program]")
   endif
_IEQuit($oIE)
_KillProcess ("iexplore.exe")
sleep (5000)
$timer=$timer+10
$oIE = _IECreate(_GetData($iteration, "URL"))
_IELoadWait($oIE)
sleep (5000)
   if WinExists("Internet Explorer","&Close program")=1 Then
	  ControlClick ("Internet Explorer","&Close program", "[CLASS:Button;Text:&Close program]")
	  ControlClick ("Internet Explorer","&Close program", "[CLASS:Button;Text:&Close program]")
   endif
until WinExists("Insert title here","")=1 or $timer>240
_ResultReporter ("Wait till the page is fully loaded",0)
_IELoadWait($oIE)
sleep(4000)
_ResultReporter ("Send Keystokes to find the Hello world message",0)
_WinWaitActiveMod("Insert title here","",60)
Send ("{CTRLDOWN}")
Send ("f")
Send ("{CTRLUP}")
sleep(4000)
_WinWaitActiveMod("Insert title here","",60)
ControlSetText ("Insert title here","","[CLASS:Edit;INSTANCE:2]",_GetData($iteration, "ValidationText"))
Send ("{ENTER}")
Sleep(2000)
if StringCompare(ControlGetText("Insert title here", "", "[CLASS:Static;INSTANCE:2]"),"1 match")=0 then
_ResultReporter ("Test Passed",1)
;msgbox (0,"Test Result", "Test Passed")
Else
_ResultReporter ("Test Failed",2)
;msgbox (0,"Test Result", "Test Falied")
EndIf
_ResultReporter ("Close IE",0)
_IEQuit($oIE)
_KillProcess ("iexplore.exe")
_KillProcess ("iexplore.exe")
if WinExists("Internet Explorer","&Close program")=1 Then
ControlClick ("Internet Explorer","&Close program", "[CLASS:Button;Text:&Close program]")
ControlClick ("Internet Explorer","&Close program", "[CLASS:Button;Text:&Close program]")
endif
_KillProcess ("csmonitor.exe")
_ResultReporter ("Close csmonitor.exe",0)
_KillProcess ("dFmonitor.exe")
_ResultReporter ("Close dFmonitor.exe",0)
_KillProcess("DFService.exe")
_ResultReporter ("Close DFService.exe",0)
_KillProcess ("DFUI.exe")
_ResultReporter ("Close DFUI.exe",0)
_KillProcess ("cscript.exe")
_ResultReporter ("Close cscript.exe",0)
_KillProcess ("DSServiceLDB.exe")
_KillProcess ("WaHostBootstrapper.exe")
_KillProcess ("WaWorkerHost.exe")
_KillProcess ("WAStorageEmulator.exe")
_ResultReporter ("Close DSService.exe",0)
_KillProcess ("Delete the Azure Project")
Sleep(10000)
_WinWaitActiveMod("Java EE ","",60)
ControlClick ("Java EE", "","[CLASS:SysTreeView32;INSTANCE:1]","left",1, 30, 10)
sleep(1000)
Send ("{DELETE}")
Sleep(2000)
_WinWaitActiveMod("Delete Resources","",60)
ControlCommand("Delete Resources","","[CLASS:Button;INSTANCE:1]","Check")
ControlClick("Delete Resources","","[CLASS:Button;Text:OK]")
sleep(2000)
if WinExists("Delete Resources", "Review the information") = 1 Then
   ControlClick("Delete Resources", "Review the information","[CLASS:Button;Text:Con&tinue]")
EndIf
WinWaitClose ("Delete Resources","",300)
sleep(5000)

_ResultReporter ("Delete the Dynamic web project",0)
_WinWaitActiveMod("Java EE ","",60)
ControlClick ("Java EE", "","[CLASS:SysTreeView32;INSTANCE:1]","left",1, 30, 10)
ControlClick ("Java EE", "","[CLASS:SysTreeView32;INSTANCE:1]","right",1, 30, 10)
Send ("d")
Send ("{ENTER}")
Sleep(2000)
_WinWaitActiveMod("Delete Resources","",60)
ControlCommand("Delete Resources","","[CLASS:Button;INSTANCE:1]","Check")
ControlClick("Delete Resources","","[CLASS:Button;Text:OK]")
WinWaitClose ("Delete Resources","",300)
sleep(2000)
_ResultReporter ("Close Eclipse",0)
WinActivate("Java EE ","")
Send ("!f")
Send ("x")

;TearDown ******************************************************************************************************************************************
_tearDown()




