		importwin32com.clientaswin32
		excel=win32.gencache.EnsureDispatch('Excel.Application')
		excel.Visible=True
		wb=excel.Workbooks.Add()
		ws=wb.Workbooks("Sheet1")

