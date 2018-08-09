Dim StartTime,EndTime
StartTime = Timer
EndTime= Timer
TimeTakem=EndTime-StartTime
print TimeTaken
'Main Function for opening the Application ------------------------

Function StartApplication
On Error resume Next
SystemUtil.Run "C:\maspects\iRISupply\iRISupply\iRISupply.exe"
wait(4)
WpfWindow("SplashScreen").Exist
Set PinXL = CreateObject("Excel.application")
Set PinWB = PinXL.Workbooks.Open("C:\maspects\Trial.xls")                          '// Login and Enable Debug Window in application
Set PinWS = PinWB.Worksheets("Sheet1")
PinXL.Visible = True
 If WpfWindow("SplashScreen").Exist  =  "True"  Then
	StartApplication = "Passed"
	Else
	StartApplication = "Failed"
End If

PinWB.Close
PinXL.Quit
End Function
'************Driver Script*******************
Call DRIVER
EndTime = Timer
TimeTaken = EndTime - StartTime
print TimeTaken
'msgbox TimeTaken
'************Function to Create Excel Connection**********************************
Function DRIVER
On Error Resume Next

Dim StartTime,EndTime'
StartTime = Timer

Set myxl = CreateObject("Excel.application")

myxl.Workbooks.Open "C:\maspects\Trial.xls" 
myxl.Application.Visible = true
 
'this is the name of  Sheet  in Excel file "qtp.xls"   where data needs to be entered 
set mysheet = myxl.ActiveWorkbook.Worksheets("Sheet1")

Row_Count=mysheet.UsedRange.Rows.Count
'msgbox Row_Count
Col_Count=mysheet.UsedRange.Columns.Count
'msgbox Col_Count

For i=2 to Row_Count

                 keyword="none" 
			If mysheet.cells(i,5).Value="Yes"  Then

                  keyword=mysheet.cells(i,4).Value
                   
			  'msgbox keyword
		    End If 
'  ----------------------Select_Case Functions are from here-------------------------
Select Case keyword
Case "StartApplication"
mysheet.cells(i,7).Value= StartApplication
Case "Login"
mysheet.cells(i,7).Value= Login
Case "Select_patient"
mysheet.cells(i,7).Value= Select_patient
Case "DebugWindow"
mysheet.cells(i,7).Value= DebugWindow
Case "RFID_ASSOCIATE"
mysheet.cells(i,7).Value= RFID_ASSOCIATE

End Select

Next
myxl.ActiveWorkbook.Save

myxl.ActiveWorkbook.Close

myxl.Application.Quit

        Set mysheet =nothing

        Set myxl = nothing

End Function

'************Function to Login**********************************
Function Login
On Error Resume Next
wait(1)
WpfWindow("Debug Window").Minimize
WpfWindow("SplashScreen").Click
wait(2)
Set PinXL = CreateObject("Excel.application")
Set PinWB = PinXL.Workbooks.Open("C:\maspects\Trial.xls")
Set PinWS = PinWB.Worksheets("Sheet2")

PinXL.Visible = False

Res1 = PinWS.cells(2,2).value
Res2 = PinWS.cells(2,3).value
'                msgbox Res1
'				msgbox Res2

varr1 = Split(Res1," ")

    For i = 0 to Ubound(varr1)
'		                 msgbox varr1(i)

	        	 Input = varr1(i)
'		                 msgbox input

         Select Case Input

               Case "0"
				  WpfWindow("Numeric").WpfButton("0").Click

	     	  Case "1"
				 WpfWindow("Numeric").WpfButton("1").Click

              Case "2"
				 WpfWindow("Numeric").WpfButton("2").Click

			  Case "3"
				 WpfWindow("Numeric").WpfButton("3").Click

			  Case "4"
				 WpfWindow("Numeric").WpfButton("4").Click

			  Case "5"
				  WpfWindow("Numeric").WpfButton("5").Click

			  Case "6"
				  WpfWindow("Numeric").WpfButton("6").Click

			   Case "7"
				   WpfWindow("Numeric").WpfButton("7").Click

               Case "8"
					WpfWindow("Numeric").WpfButton("8").Click

			  Case "9"
					WpfWindow("Numeric").WpfButton("9").Click
		     End Select
	Next
	
	
WpfWindow("Numeric").WpfButton("Enter").Click
wait(2)

If WpfWindow("iRIScope").WpfObject("Please Select a Patient").Exist="True" Then
	Login="Passed"
	else
	Login="Failed" 
   End If

PinWB.Close
PinXL.Quit

 Wait (3)
  
End Function
'-----------Select_patient----------
Function Select_patient
 On Error Resume Next

WpfWindow("iRIScope").WpfTable("dgrPatientList").SelectCell 7,"MRN" @@ hightlight id_;_1971285312_;_script infofile_;_ZIP::ssf376.xml_;_
WpfWindow("iRIScope").WpfButton("Confirm").Click

wait(2)
  If WpfWindow("iRIScope").WpfObject("Access Cabinet").Exist="True"  Then
     Select_patient="Passed"
	 else
	 Select_patient="Failed" 
     End If
     
WpfWindow("iRIScope").WpfButton("Logout").Click @@ hightlight id_;_1971280560_;_script infofile_;_ZIP::ssf392.xml_;_
PinWB.Close
PinXL.Quit
 Wait (3)
 Call Logout
  
End Function

'-----------Logout----------
Function Logout
On Error Resume Next
wait(1)
WpfWindow("SplashScreen").Click
wait(2)
Set PinXL = CreateObject("Excel.application")
Set PinWB = PinXL.Workbooks.Open("C:\maspects\Trial.xls")
Set PinWS = PinWB.Worksheets("Sheet2")

PinXL.Visible = False

Res1 = PinWS.cells(8,2).value
Res2 = PinWS.cells(2,3).value
'                msgbox Res1
'				msgbox Res2

varr1 = Split(Res1," ")

    For i = 0 to Ubound(varr1)
'		                 msgbox varr1(i)

	        	 Input = varr1(i)
'		                 msgbox input

         Select Case Input

               Case "0"
				  WpfWindow("Numeric").WpfButton("0").Click

	     	  Case "1"
				 WpfWindow("Numeric").WpfButton("1").Click

              Case "2"
				 WpfWindow("Numeric").WpfButton("2").Click

			  Case "3"
				 WpfWindow("Numeric").WpfButton("3").Click

			  Case "4"
				 WpfWindow("Numeric").WpfButton("4").Click

			  Case "5"
				  WpfWindow("Numeric").WpfButton("5").Click

			  Case "6"
				  WpfWindow("Numeric").WpfButton("6").Click

			   Case "7"
				   WpfWindow("Numeric").WpfButton("7").Click

               Case "8"
					WpfWindow("Numeric").WpfButton("8").Click

			  Case "9"
					WpfWindow("Numeric").WpfButton("9").Click
		     End Select
	Next
WpfWindow("Numeric").WpfButton("Enter").Click

wait 1
PinWB.Close
PinXL.Quit

 Wait (3)
 End Function
 '-----------LaunchApplication----------
 Function LaunchApplication
On Error resume Next
SystemUtil.Run "C:\maspects\iRISupply\iRISupply\iRISupply.exe" @@ hightlight id_;_1966628_;_script infofile_;_ZIP::ssf413.xml_;_
WpfWindow("Debug Window").Minimize
 @@ hightlight id_;_918120_;_script infofile_;_ZIP::ssf455.xml_;_
 @@ hightlight id_;_2164094_;_script infofile_;_ZIP::ssf414.xml_;_
Wait (1)
 End Function
 '************Function to ApplicationLogin**********************************
Function ApplicationLogin
On Error Resume Next
wait(1)

 @@ hightlight id_;_2230066_;_script infofile_;_ZIP::ssf456.xml_;_
WpfWindow("SplashScreen").Click
wait(2)
Set PinXL = CreateObject("Excel.application")
Set PinWB = PinXL.Workbooks.Open("C:\maspects\Trial.xls")
Set PinWS = PinWB.Worksheets("Sheet2")

PinXL.Visible = False

Res1 = PinWS.cells(2,2).value
Res2 = PinWS.cells(2,3).value
'                msgbox Res1
'				msgbox Res2

varr1 = Split(Res1," ")

    For i = 0 to Ubound(varr1)
'		                 msgbox varr1(i)

	        	 Input = varr1(i)
'		                 msgbox input

         Select Case Input

               Case "0"
				  WpfWindow("Numeric").WpfButton("0").Click

	     	  Case "1"
				 WpfWindow("Numeric").WpfButton("1").Click

              Case "2"
				 WpfWindow("Numeric").WpfButton("2").Click

			  Case "3"
				 WpfWindow("Numeric").WpfButton("3").Click

			  Case "4"
				 WpfWindow("Numeric").WpfButton("4").Click

			  Case "5"
				  WpfWindow("Numeric").WpfButton("5").Click

			  Case "6"
				  WpfWindow("Numeric").WpfButton("6").Click

			   Case "7"
				   WpfWindow("Numeric").WpfButton("7").Click

               Case "8"
					WpfWindow("Numeric").WpfButton("8").Click

			  Case "9"
					WpfWindow("Numeric").WpfButton("9").Click
		     End Select
	Next
	
	
WpfWindow("Numeric").WpfButton("Enter").Click

PinWB.Close
PinXL.Quit

 Wait (3)
  
End Function

'-----------DebugWindow----------

Function DebugWindow
On Error resume Next
Call LaunchApplication
wait (2)
Call ApplicationLogin
wait (2)
Set PinXL = CreateObject("Excel.application")
Set PinWB = PinXL.Workbooks.Open("C:\maspects\Trial.xls")
Set PinWS = PinWB.Worksheets("Sheet2")

PinXL.Visible = False

Res1 = PinWS.cells(9,2).value
WpfWindow("Debug Window").MakeVisible
WpfWindow("Debug Window").WpfEdit("textBox1").Set "78024511"

WpfWindow("Debug Window").WpfButton("Send").Click @@ hightlight id_;_1916159616_;_script infofile_;_ZIP::ssf447.xml_;_
 @@ hightlight id_;_4523794_;_script infofile_;_ZIP::ssf448.xml_;_
WpfWindow("iRIScope").WpfTable("dgrPatientList").SelectCell 0,"MRN"

WpfWindow("iRIScope").WpfTable("dgrPatientList").SelectCell 0,"MRN" @@ hightlight id_;_1979075368_;_script infofile_;_ZIP::ssf453.xml_;_
WpfWindow("Debug Window").Minimize @@ hightlight id_;_131894_;_script infofile_;_ZIP::ssf454.xml_;_
 @@ hightlight id_;_2033966064_;_script infofile_;_ZIP::ssf464.xml_;_

WpfWindow("iRIScope").WpfButton("Confirm").Click

If WpfWindow("iRIScope").WpfObject("Access Cabinet").Exist="True"  Then
     DebugWindow="Passed"
	 else
	 DebugWindow="Failed" 
     End If
     
WpfWindow("iRIScope").WpfButton("Logout").Click @@ hightlight id_;_1971280560_;_script infofile_;_ZIP::ssf392.xml_;_
PinWB.Close
PinXL.Quit
 Wait (3)
 Call Logout

End Function
'--------------------RFID_ASSOCIATE---------------------------------------'
Function RFID_ASSOCIATE
Dim con
Dim rs,strSQL
Set con=createobject("adodb.connection")
Set rs=Createobject("adodb.recordset")
Set PinXL = CreateObject("Excel.application")
Set PinWB = PinXL.Workbooks.Open ("C:\Maspects\Trial.xls" )
set pinWS = PinWB.Worksheets("SQL_Connection")
varr = Cstr(pinWS.Cells(2,1).Value)
varr1 =  trim(varr)
usename =  Cstr(pinWS.Cells(2,2).Value)
UN = trim(usename)
password = Cstr(pinWS.Cells(2,3).Value)
PWD = trim(password)
IrisDB = Cstr(pinWS.Cells(2,4).Value)
DB = trim(IrisDB)
con.open"provider=sqloledb.1;server=" & varr1 & ";uid=" & UN & ";pwd=" & PWD & ";database=" & DB &""
rs.open  "Select  Top 1 *  from tblproducts where ProducttypeID = 1",con
ProductID = rs.Fields("ProductID")
VarrProdID =  ProductID
print VarrProdID
Set pinWS1 = PinWB.Worksheets("RFID")
Normal_RFID1  = pinWS1.Cells(2,2).value
Normal_RFID2  = pinWS1.Cells(2,3).value
Normal_RFID3  = pinWS1.Cells(2,4).value
Expired_RFID1 = pinWS1.Cells(3,2).value
print Normal_RFID1
print Normal_RFID2
print Normal_RFID3
print Expired_RFID1
Con.Execute ( "insert into TBLITEMS (ProductID,RFID,[Expired Date],ItemStatusID) values (" & VarrProdID & ",'" & pinWS1.Cells(2,2).value & "','03-19-2019',0)")    '// Normal RFID 1
Con.Execute ( "insert into TBLITEMS (ProductID,RFID,[Expired Date],ItemStatusID) values (" & VarrProdID & ",'" & pinWS1.Cells(2,3).value & "','03-19-2019',0)")   '// Normal RFID 2
Con.Execute ( "insert into TBLITEMS (ProductID,RFID,[Expired Date],ItemStatusID) values (" & VarrProdID & ",'" & pinWS1.Cells(2,4).value & "','03-19-2019',0)")   '// Normal RFID 2
Con.Execute ( "insert into TBLITEMS (ProductID,RFID,[Expired Date],ItemStatusID) values (" & VarrProdID & ",'" & pinWS1.Cells(3,2).value & "','03-19-2010',0)")      '// Expired RFID`
strSQL="Select  Top 1 RFID  from TBLITEMS order by LastActivityDate desc"
Set rs=con.execute(strSQL)
ITEM=rs.Fields("RFID")

If ITEM=trim(Expired_RFID1) Then

	RFID_ASSOCIATE="Passed"
	 else
	 RFID_ASSOCIATE="Failed" 
End If
PinWB.Close
PinXL.Quit
End Function







 @@ hightlight id_;_1971282240_;_script infofile_;_ZIP::ssf379.xml_;_
