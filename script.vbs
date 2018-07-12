'usage  script.vbs 1dc-lan-userid 1dc-lan-password JIRA-SDLC-FilterID
'usage  script.vbs mukherri n0tmyp@$$w0rd 18317 dsr_xlsm

Option Explicit
Dim User_ID,User_Pass,Filter_ID
Dim File_Save_Path,File_Save_Name,Download_File,URL,url_POST,url_GET,POST,RequestHeader,ContentType,staticEXTNhtml,staticEXTNxls,staticEXTNxlsm
Dim nTables, table, TableData, Elem, Count_Table, Count_Table_Row, Count_Table_Header
Dim fileExcel, objExcel, strExcelPath, objSheet
dim i,j
Wsh.Echo ("Web File Processing")
User_ID=WScript.Arguments.Item(0):User_Pass=WScript.Arguments.Item(1):Filter_ID=WScript.Arguments.Item(2):fileExcel=WScript.Arguments.Item(3)
staticEXTNhtml=".html":staticEXTNxls=".xls":staticEXTNxlsm=".xlsm"
Dim scriptdir:scriptdir=CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
File_Save_Path=scriptdir+"\":File_Save_Name=Filter_ID+"_"+Replace(CStr(Timer),".","_")+staticEXTNhtml:
Download_File=File_Save_Path+File_Save_Name
strExcelPath = File_Save_Path+"dsr_xlsm"+staticEXTNxlsm:URL="https://jirasdlc.1dc.com":url_POST=URL+"/login.jsp"
url_GET=URL+":8443/sr/jira.issueviews:searchrequest-excel-current-fields/"+Filter_ID+"/SearchRequest-"+Filter_ID+".xls"
POST="os_username="+User_ID+"&os_password="+User_Pass+"&os_destination=&user_role=&atl_token=&login=Log+In"
RequestHeader="":ContentType="application/x-www-form-urlencoded"
Set objExcel = CreateObject("Excel.Application"):objExcel.WorkBooks.Open strExcelPath:Set objSheet = objExcel.ActiveWorkbook.Worksheets("general_report")
strExcelPath = File_Save_Path+"dsr_xlsm"+staticEXTNxlsm
For i = objSheet.UsedRange.Row+3 to objSheet.UsedRange.Rows.Count
	For j = objSheet.UsedRange.Column to objSheet.UsedRange.Columns.Count
		objSheet.Cells(i, j).Value = ""
	Next
Next
dim xHttp:Set xHttp=createobject("Microsoft.XMLHTTP")
With xHttp
   .Open"POST",url_POST,False:.setRequestHeader"Content-Type",ContentType:.send POST:.Open "GET",url_GET,False:.Send RequestHeader
End With
dim bStrm:Set bStrm=createobject("Adodb.Stream")
with bStrm:.type=1:.open:.write xHttp.responseBody:.savetofile Download_File, 2:end with:Stop
Set xHttp=Nothing:Set bStrm=Nothing:Set scriptdir=Nothing
Dim ie:Set ie=CreateObject("InternetExplorer.Application")
With ie:.Navigate Download_File:Do until .ReadyState=4:WScript.Sleep 50:Loop
	With .document:Dim theTables:set theTables=.all.tags("table"):nTables = theTables.length
		for each table in theTables
			Count_Table=Count_Table+1:Count_Table_Row=0:Count_Table_Header=0
			For Each Elem in table.GetElementsByTagName("TR"):Count_Table_Row=Count_Table_Row+1:Next
			For Each Elem in table.GetElementsByTagName("TH"):Count_Table_Header=Count_Table_Header+1:Next
			if(Count_Table_Header<>0)then
				for i=0to Count_Table_Row-1
					for j=0 to Count_Table_Header-1
						objSheet.Cells(i+4, j+1).Value = table.rows(i).cells(j).innerText
					next
					TableData=TableData&vbNewLine&vbNewLine
				next
			end if
		next
	End With
End With
objExcel.ActiveWorkbook.Save:objExcel.ActiveWorkbook.Close:objExcel.Application.Quit
Set ie=Nothing:Set theTables=Nothing:Set objSheet=Nothing: Set objExcel=Nothing
Wsh.Echo ("Web File Processed")
