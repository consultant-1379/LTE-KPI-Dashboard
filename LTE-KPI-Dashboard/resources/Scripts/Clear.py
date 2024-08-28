Document.Properties["KPIList"] = ""
Document.Properties["SelectionWarning"] = ""

for i in range(1,13):
	Document.Properties["SelectedKPI"+str(i)] = ""
	Document.Properties["SelectedKPIName"+str(i)] = ""
	Document.Properties["Units"+str(i)] = ""
	Document.Properties["Note"+str(i)] = ""

