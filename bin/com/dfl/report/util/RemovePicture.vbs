
Dim debug, devTest
debug = False
devTest = False

If Not debug Then
 On Error Resume Next
End If

	If WScript.Arguments.Count < 1 Then
		MsgBox "Usage: SubstMacros-Excel <xls file> "
		WScript.Quit
	End If

	Dim xlsFileName
	xlsFileName = WScript.Arguments(0)
	'filePath = WScript.Arguments(1)
	'sheetname = WScript.Arguments(1)

	Dim fs, txtFile, line, data
	Set fs = CreateObject("Scripting.FileSystemObject")
	If fs.FileExists(xlsFileName) = False Then
	  WScript.Echo "File " & xlsFileName & " doesn't exist!"
	  WScript.Quit
	End If
	   
	Dim excelApp
	Set excelApp = CreateObject("Excel.Application")
	excelApp.Visible = false
	'excelApp.Visible = true
	'excelApp.ScreenUpdating = debug

	Dim wb
	Set wb = excelApp.Workbooks.Open(xlsFileName)
	
	
   Dim sht,P,shname,l1,t1,l2,t2,l11,t11,l21,t21,l12,t12,l22,t22,l13,t13,l23,t23,l14,t14,l24,t24,l15,t15,l25,t25
   For Each sht In wb.Sheets
   		sht.Activate
		shname = Mid(sht.Name,1,5)
		'MsgBox shname
		If shname = "4_部品构" Or shname = "14_ST" Or shname = "8_Loc" Then 
		l1 = sht.cells(3,2).Left
		t1 = sht.cells(3,2).Top
		l2 = sht.cells(47,62).Left
		t2 = sht.cells(47,62).Top
			For Each P In sht.Shapes
				If p.Left>=l1 And p.Left<l2 And p.Top>=t1 And p.Top<t2 Then
				   p.Delete
                END If	
             Next 
        End If	 
		
		If shname = "9_断面定"  Then 
		    l1 = sht.cells(5,21).Left
		    t1 = sht.cells(5,21).Top
		    l2 = sht.cells(16,36).Left
		    t2 = sht.cells(16,36).Top
			
			l11 = sht.cells(22,21).Left
		    t11 = sht.cells(22,21).Top
		    l21 = sht.cells(33,36).Left
		    t21 = sht.cells(33,36).Top
			
			l12 = sht.cells(39,21).Left
		    t12 = sht.cells(39,21).Top
		    l22 = sht.cells(50,36).Left
		    t22 = sht.cells(50,36).Top
			
			l13 = sht.cells(5,60).Left
		    t13 = sht.cells(5,60).Top
		    l23 = sht.cells(16,75).Left
		    t23 = sht.cells(16,75).Top
			
			l14 = sht.cells(22,60).Left
		    t14 = sht.cells(22,60).Top
		    l24 = sht.cells(33,75).Left
		    t24 = sht.cells(33,75).Top
			
			l15 = sht.cells(39,60).Left
		    t15 = sht.cells(39,60).Top
		    l25 = sht.cells(50,75).Left
		    t25 = sht.cells(50,75).Top
						
			
			For Each P In sht.Shapes			
				If p.Left>=l1 And p.Left<l2 And p.Top>=t1 And p.Top<t2 Then
				   p.Delete
                END If	
				If p.Left>=l11 And p.Left<l21 And p.Top>=t11 And p.Top<t21 Then
				   p.Delete
                END If	
				If p.Left>=l12 And p.Left<l22 And p.Top>=t12 And p.Top<t22 Then
				   p.Delete
                END If	
				If p.Left>=l13 And p.Left<l23 And p.Top>=t13 And p.Top<t23 Then
				   p.Delete
                END If	
				If p.Left>=l14 And p.Left<l24 And p.Top>=t14 And p.Top<t24 Then
				   p.Delete
                END If	
				If p.Left>=l15 And p.Left<l25 And p.Top>=t15 And p.Top<t25 Then
				   p.Delete
                END If	
             Next 
        End If	
		
	Next
	
	wb.Save=True
	wb.Close 0
	excelApp.Quit 0
	WScript.Quit
