<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="check_user_inc.asp"-->
<!--#INCLUDE FILE="general_text.asp"--> <!-- My external error writing file -->
<%

on error goto 0
'on error resume next

server.scripttimeout = 36000

err = 0
sesxl = False
validdb = False
dropxl = false

listID = request("ULREF")
SESSION("listID") = listID


errorline=1
if err > 0 then	someerror

' Error Checking written to log file -- Verify listID value
elog.writeline("List ID:" & listID)

hua = UCASE(Request.ServerVariables("HTTP_USER_AGENT"))
if instr(hua,"MAC") > 0 then isMAC = true

count = 1

IF count = 1 THEN

'-------------------------------------------------
    'General Routine for list upload process
'-------------------------------------------------

        if listid = -1 then
				'make up a test filename
                filename = "XXXtest.csv"
                session("dbfilename") = filename
        else
				Set cn = Server.CreateObject("ADODB.Connection")
				Set rs = Server.CreateObject("ADODB.Recordset")
				cn.ConnectionTimeout = 15
				cn.CommandTimeout = 180
				'cn.Open "DSN=xABSMAIN"
				CN.Provider = "MSDASQL"
				CN.Open "Driver={SQL Server};Server=DBServe;UID=absgen;PWD=absgencb"

				'Perhaps there is no valid filename in this table?
	        	set rs = cn.execute("select * from FFTWuploads where uploadID = " & listid)

	        	filename = rs("fileName")

				rs.close
	       		cn.close

	       		session("dbfilename") = filename

        		set rs = nothing
	        	set cn = nothing
        end if


        SET ab = Server.CreateObject("ABSFile.Utility")

		oldpath = "d:\webroot\upload\" & filename
		'Creates file db*****.fil
		NEWFILENAME = "DB" & listid & ".fil"
		
  		path = "d:\webroot\upload\db\" & NEWFILENAME
		xpath = "\\192.168.1.123\d\webroot\upload\db\" & NEWFILENAME

		res = ab.copyfile(oldpath,path)
    	'Copying file but not deleting old file from d:\webroot\upload\ // Comments Added by JTD 7/26/05
        if listid <> -1 then res = ab.deletefile(oldpath)

        set ab = nothing

		SESSION("DBPATH") = path
		filename = NEWFILENAME
		'Write File path to Error Log
		elog.writeline("Filename:" & filename)
		elog.writeline("File Path:" & path)
		
		xlflg = 0
		
		
		'Check File Type
		set ft = server.createobject("FileType.CHECKFILE")
        id = ft.id(path)
        set ft = nothing
	    elog.writeline("File Type ID:" & id)
	
'-------------------------------------------------
    'Subroutine for Excel Doc Process
'-------------------------------------------------
	if id = 12 or id = 13 or id = 14 or id = 15 then

        	validdb = true

       	 	SESSION("IMPDBTP") = "Excel"
			sesxl = True

			xlflg = 1
            procxl = 0
            Set xl = Server.CreateObject("ExcelServ.XL")
			dim tsheetscount
			dim res

			tsheetscount = 0
			xl.visible = false

            fpath = path
			res = xl.check(xpath,tsheetscount,sheetname,sheetcount)
            validsheets = 0


elog.writeline("res:" & res)
elog.writeline("tsheetscount:" & tsheetscount)
' need these loops to start counting from 0
            if res = 0 then
					for i = 1 to tsheetscount
                		if sheetcount(i) > 0 then validsheets = validsheets + 1 : lastvalidsheet = i
					next
					elog.writeline("validsheets:" & validsheets)
                    if validsheets > 1 then
						response.write("&XTSC=" & tsheetscount) 'No server encoding gives me a better response into flash
						elog.writeline("&XTSC=" & tsheetscount) 
						for i = 1 to tsheetscount
	                    	response.write("&XSN" & i-1 & "=" & server.urlencode(sheetname(i)))
							response.write("&XSC" & i-1 & "=" & sheetcount(i))

							elog.writeline("&XSN="& i-1 & "=" & sheetname(i))
							elog.writeline("&XSC="& i-1 & "=" & sheetcount(i))
						next
                        SESSION("XLPATH") = xpath
                    else
						dropxl = true
						response.write("&XTSC=" & "0")

						sheetn = lastvalidsheet
                        coutfile = "\\192.168.1.122\d\webroot\upload\db\DB" & listid & ".CSV"
                        filename = "DB" & listid & ".CSV"
                        hiddenfields = 0
                        eres = xl.export(xpath, coutfile, sheetn, hiddenfields, hiddennumbers)
                        if eres > 0 then
							err = 1
                            res = 888
						else
                            response.write("&XHFDS=" & server.urlencode(hiddenfields))

							for i = 1 to hiddenfields
								response.write("&XHF" & i & "=" & server.urlencode(hiddennumbers(i)))
							next
							procxl = 1
                            xlflg = 0
                        end if
                    end if
			end if
        xl.kill
		set xl = nothing
	
	else
  		if isMAC = true then
			set fmac = server.createobject("fixMac.mac1")
			fmac.fixFile(path)
			set fmac = nothing
		end if
	end if

'-------------------------------------------------
    'Subroutine for Other Document Processes
'-------------------------------------------------

elog.writeline("procxl:" & procxl)
    if id = 0 or id = 2 or id = 3  or id = 35 or id = 36 or id = 37 or procxl = 1 then
elog.writeline("In Second If (non-excel)")

        validdb = true
		  		
		' reID File Type  
		Select Case id
        Case 2, 3
        DBT = "csv file"
    	Case 35
        DBT = "DBase Ver 3, 4"
    	Case 36
        DBT = "DBase Ver 2"
		End Select 		  
		  
		Const dbVersion10 = 1
		Const dbVersion11 = 8
		Const dbVersion20 = 16
		Const dbVersion30 = 32
		Const dbVersion40 = 64
		Sub CreateNewMDB(FileName, Format)
  		Dim Engine 
  		Set Engine = CreateObject("DAO.DBEngine.36")
  		Engine.CreateDatabase FileName, ";LANGID=0x0409;CP=1252;COUNTRY=0", Format
		End Sub

		'Create Access2000 database
		CreateNewMDB "d:\webroot\temp\db\TDB" & listid & ".MDB", dbVersion40
		cnstr = "DSN=mdbdata;DBQ=d:\webroot\temp\db\TDB" & listid & ".MDB"

		elog.writeline(cnstr)

		' --- Öpen der Database
		Set cn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")
    	cn.open cnstr

		'Create csv file path 
		Dim csv_path
		Dim csv_file
   		csv_path = "d:\webroot\upload\db\"
   		csv_file =""'nothing

   		If sesxl = True Then
   			csv_file = "DB" & listid & ".CSV"
   		errorline=69
		if err > 0 then	someerror

   		Else
   			csv_file = "DB" & listid & ".FIL"
   		End If

		'Read in the csv_file, figure out how many columns, add a row to the top of that file with that many columns so it
		'can be swallowed up by access rather than the first row of the customers data
		Dim mod_csv_file
		Dim mod_csv_file_name
		Dim in_csv_file
		Dim strText
		Dim outText
		Dim varArray
		Dim numCols
		Dim cnt

		mod_csv_file_name = "MDB" & listid & ".CSV"
		mod_csv_file = csv_path & mod_csv_file_name
		in_csv_file = csv_path & csv_file
		set fs = server.CreateObject("Scripting.FileSystemObject")
		set outFile = fs.CreateTextFile(mod_csv_file,TRUE,FALSE)
		set inFile = fs.OpenTextFile(in_csv_file, 1, False)

		If inFile.AtEndOfStream <> True then
			strText = inFile.ReadLine
			varArray = Split(strText, ",")
			numCols = UBound(varArray)
			for cnt = 1 to numCols
				outText = outText & "Col" & cnt & ", "
			next
			outText = outText & "Col" & cnt
			outFile.writeline(outText)
			outFile.writeline(strText)
		end if
		Do While inFile.AtEndOfStream <> True
			strText = inFile.ReadLine
			if Len(strText) > 2 then ' using two just in case \n\r counts as two and is in there
				outFile.writeline(strText)
			end if
		loop
		inFile.close
		outFile.close

		file_path = "[" & mod_csv_file_name & "]" 

		' --- SQL-String zum Neuerstellen einer Tabelle mit allen Daten aus der CSV-Datei und Ausführen des SQL-Statements 
		sqlstr = "SELECT * INTO DATA FROM" & file_path & "IN """ & csv_path & """ ""TEXT;"""
   		set rs = cn.execute (sqlstr)
   		elog.writeline("sql string: " & sqlstr)
				
		set rs = cn.execute ("CREATE TABLE INFO (count INTEGER, dbtype varchar)")
		set rs = cn.execute ("INSERT INTO INFO VALUES ('0', '" & DBT & "')")
				
   		' --- Terminate db connection
   		cn.Close
		Set rs = Nothing
		Set cn = Nothing
		
		
		'Error Checking lines - written to text file
		elog.writeline("List ID After ABS.dll:" & listid)
		elog.writeline("ERROR #:" & err)
		elog.writeline("Record Count:" & recordcount)
		'set dbx = nothing
	end if
		
END IF


elog.writeline("validdb:" & validdb)
elog.writeline("xlflg:" & xlflg)
IF validdb = true THEN
    	if res < 1 then err = 0

		if xlflg = 0 and res = 0 then
	   		Set cn = Server.CreateObject("ADODB.Connection")
			Set rs = Server.CreateObject("ADODB.Recordset")
			cn.ConnectionTimeout = 15
			cn.CommandTimeout = 180
			cnstr = "DSN=mdbdata;DBQ=d:\webroot\temp\db\TDB" & listid & ".MDB"
			
			'Write File path to Error Log
			elog.writeline("New File Path:" & cnstr)
			
	   		cn.Open cnstr
			set rs = cn.execute("select * from INFO")
			'Checks to see if file is excel or DB/CSV then assigns value
            if sesxl = False then SESSION("IMPDBTP") = RS(1) else SESSION("IMPDBTP") = "Excel"

		   	set rs = cn.execute("select count(*) from DATA")
			SESSION("IMPDBCT") = RS(0)

			set rs = NOTHING
	   		set cn = NOTHING
		end if

ELSE
  	 err = 99
END IF


'-------------------------------------------------
    'Spits Error Number to Browser Window
'-------------------------------------------------

' This code called in Flash MC Instance "_level0.dbup"
'response.write("DBUPERR=" & server.urlencode(err) & "&")
response.write("&DBUPERR=" & err) 
response.write("&DBDONE=1")  ' What the 'ell are these for?

'This code writes error to Text File
elog.writeline("Server Error Code:" & err)
elog.writeline(" ")
elog.writeline(" ")
'-------------------------------------------------
    'Spits Error Number to Database
'-------------------------------------------------
	
REM ERROR REPORTING CODE!!
IF err > 0 THEN
			Set ecn = Server.CreateObject("ADODB.Connection")
			Set ers = Server.CreateObject("ADODB.Recordset")
			ecn.ConnectionTimeout = 15
			ecn.CommandTimeout = 180
			'ecn.Open "DSN=xABSMAIN"
			eCN.Provider = "MSDASQL"
			eCN.Open "Driver={SQL Server};Server=DBServe;UID=absgen;PWD=absgencb"
			sql = "insert into FFTWerrors values (getdate()," & session("ID") & "," & SESSION("ContactNumber") & "," & session("JOBNUM") & ",'" & request.servervariables("PATH_INFO") & "','" & err.source & "'," & err.number & ",'" & err.description & " | " & ID & "','" & request.serverVariables("HTTP_USER_AGENT") & "'," & "'" & SESSION("NOFS") & "','" & SESSION("JAVAENABLED") & "')"
			set ers = ecn.execute(sql)
			set ers = nothing
			set ecn = nothing
END IF

REM END ERROR REPORTING CODE!

sub SomeError
      		Set ecn = Server.CreateObject("ADODB.Connection")
			Set ers = Server.CreateObject("ADODB.Recordset")
			ecn.ConnectionTimeout = 15
			ecn.CommandTimeout = 180
			'ecn.Open "DSN=xABSMAIN"
	eCN.Provider = "MSDASQL"
	eCN.Open "Driver={SQL Server};Server=DBServe;UID=absgen;PWD=absgencb"
	
	sql = "insert into FFTWerrors values (getdate()," & session("ID") & "," 
	sql=sql & SESSION("ContactNumber") & "," & session("JOBNUM") & ",'" 
	sql=sql & request.servervariables("PATH_INFO") & "','" & err.source & "'," 
	sql=sql & err.number & ",'" & err.description & " | "& errorlog & " | "&  errorline & "','" 
	sql=sql & request.serverVariables("HTTP_USER_AGENT") & "'," 
	sql=sql & "'" & SESSION("NOFS") & "','" & SESSION("JAVAENABLED") & "')"


			set ers = ecn.execute(sql)
			set ers = nothing
			set ecn = nothing
	err.Clear
end Sub

%>
