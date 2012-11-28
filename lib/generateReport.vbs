'
' ReportGenerator Version 0.2.1
'
Class GenerateReportClass

	Public LAYOUTBOOK
	Public NEWBOOK

	Public DATXML
	Public MAPXML
	Public OBJXLS
	Public OBJWSH

	Public xlsPath
	Public mapPath
	Public datPath

	Public autoClose

	Function nvl(val1,val2)
		Dim ret
		ret = val1
		if IsNull(val1) then
			ret = val2
		end if
		nvl = ret
	End Function

	Function hasSheetName(book,name)
		Dim i,find
		find = false
		For i = 1 to book.sheets.count
			If book.sheets.item(i).name = name then
				find = true
			End if
		Next
		hasSheetName = find
	End Function

	Private Sub Class_Initialize()
		set DATXML	= CreateObject("Msxml2.DOMDocument") 
		set MAPXML	= CreateObject("Msxml2.DOMDocument") 
		set OBJXLS	= CreateObject("Excel.Application") 
		Set OBJWSH	= CreateObject("WScript.Shell")
		autoClose = true
	End Sub

	Sub GenerateReport()
		Dim objPage,objMap
		Dim objDatSheet,objCsrSheet,objMapSheet,objNode
		Dim sheetName,newSheetName
		Dim wnY,wnX,numPos,orgSheetCnt
		Dim i,j,n,v

		OBJXLS.Visible	=	true
		OBJWSH.AppActivate OBJXLS

		set LAYOUTBOOK	=	OBJXLS.workbooks.open(xlsPath)
		OBJXLS.Visible	=	false
		OBJXLS.Visible	=	true


		set NEWBOOK	=	OBJXLS.Workbooks.Add
		NEWBOOK.Application.DisplayAlerts = False

		orgSheetCnt = NEWBOOK.Sheets.count
		For i = 1 to orgSheetCnt-1
			NEWBOOK.Sheets(1).delete
		Next

		DATXML.async=False
		DATXML.Load(datPath)
		set objPage = DATXML.documentElement.getElementsByTagName("page")

		MAPXML.async=False
		MAPXML.Load(mapPath)
		set objMap = MAPXML.documentElement

		For i=0 To objPage.Length-1
			set objDatSheet = objPage(i).getElementsByTagName("sheet")
			For j=0 To objDatSheet.Length-1
				With NEWBOOK
					sheetName = objDatSheet(j).getAttribute("name")
					LAYOUTBOOK.Sheets(sheetName).Copy .Sheets(.Sheets.count)
					set objCsrSheet	= .Sheets(.Sheets.count-1)
					newSheetName = objDatSheet(j).getAttribute("newName")
					If not IsNull(newSheetName) and not hasSheetName(NEWBOOK,newSheetName) then
						objCsrSheet.name = newSheetName
					End If
				End With
				set objMapSheet = objMap.getElementsByTagName("sheet")
				For m=0 To objMapSheet.Length-1
					if objMapSheet(m).getAttribute("name") = sheetName then
						With objMapSheet(m).childNodes
				  			For n=0 To .Length-1
								set objNode = objDatSheet(j).getElementsByTagName(.item(n).nodeName)
								if objNode.Length > 0 then
									numPos = instr(.item(n).text,",")
									For v=0 To objNode.Length-1
										wnY = Cint(mid(.item(n).text,1,numPos-1))
										wnX = Cint(mid(.item(n).text,numPos+1))
										wnY = wnY + Cint(nvl(objNode(v).getAttribute("addY"),"0"))
										wnX = wnX + Cint(nvl(objNode(v).getAttribute("addX"),"0"))
										objCsrSheet.cells(wnY,wnX) = objNode(v).text
									Next
								End if
							Next
						End With
					end if
				Next
			Next
		Next

		if autoClose = true then
			LAYOUTBOOK.close
			set LAYOUTBOOK = Nothing
		end if

		With NEWBOOK
			.Sheets(.Sheets.count).delete
			.Sheets(1).select
		End With

		NEWBOOK.Application.DisplayAlerts = True

	End Sub

End Class

function generateReport(xlsPath, mapPath, datPath)
	Dim GR
	set GR = New GenerateReportClass
	GR.xlsPath = xlsPath
	GR.mapPath = mapPath
	GR.datPath = datPath
	GR.GenerateReport()
	set generateReport = GR
End Function

