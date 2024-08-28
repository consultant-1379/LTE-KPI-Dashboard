from Spotfire.Dxp.Application.Visuals import Visualization,LineChart,TablePlot, TablePlotColumnSortMode
from Spotfire.Dxp.Application.Visuals import HtmlTextArea
import re

Flag_variable = "EUtranCell"
defaultcol = ["DATETIME_ID","ERBS","EUtranCell"]
cellname = ["DATETIME_ID","ERBS","TermPointToSGW"]
Eutrancellfdd = list(Document.Properties["Eutrancellfdd"].split(":"))
Eutrancellfddvector = list(Document.Properties["Eutrancellfddvector"].split(":"))
fdduniontdd = list(Document.Properties["fdduniontdd"].split(":"))
dodleveltables= list(Document.Properties["dodleveltables"].split(":"))
Termpoint = list(Document.Properties["Termpoint"].split(":"))
Meankpis = list(Document.Properties["Meankpis"].split(":"))
Multitable_kpis = list(Document.Properties["Multitablekpis"].split(":"))  
Termpointsgw = "<[TermPointToSGW]>"
Eutrancellaxis = "<[EUtranCell]>"
Fdduniontddtable = "FDD_Union_TDD_DOD_KPI's_Nodeview"
kpiFormula = Document.Properties[Document.Properties["AccessibilityKPIs"][2:-1]]	
columnName = re.findall(r'\[([^[]+)\]', kpiFormula)
columnName = list(set(columnName))
KPInode = "KPI View - Node"
KPIcell = "KPI View - Cell"
cellview = "Cell View"
upto_7days = "Up to 7 Days"

def getdatatable(tablename):
	try:
		return Document.Data.Tables[tablename]
	except KeyError:
		raise ("Error - cannot find data table: " + tablename)

def addcolumns(myvis,columnname,defaultcol):
		datatable = myvis.Data.DataTableReference
		myvis.SortInfos.Clear()
		for item in defaultcol:
			myvis.TableColumns.Add(datatable.Columns.Item[item])
		if Document.Properties["DisplayName"] in Meankpis:
			myvis.TableColumns.Add(datatable.Columns.Item["DCVECTOR_INDEX"])
		myVis.TableColumns.Add(datatable.Columns.Item["KPI Value"])
		for item in columnname:
			myvis.TableColumns.Add(datatable.Columns.Item[item])
		myvis.SortInfos.Add(datatable.Columns["DATETIME_ID"],TablePlotColumnSortMode.Ascending)
		if Document.Properties["DisplayName"] == "Mean UL PDCP UE Throughput":
			myvis.Data.WhereClauseExpression = "[DCVECTOR_INDEX]<=26 and [KPI Value] is not null"
		else:
			myvis.Data.WhereClauseExpression = "[KPI Value] is not null"

def addaxistochart(vistitle,pagetilte,vis,flagvariable):
	if pagetilte == KPIcell:
		if  cellview  in vistitle:
			if flagvariable == "EUtranCell":
				vis.XAxis.Expression = Eutrancellaxis
				vis.ColorAxis.Expression = Eutrancellaxis
			else:
				vis.XAxis.Expression = Termpointsgw                     
				vis.ColorAxis.Expression = Termpointsgw
		elif  upto_7days in vistitle:
			if flagvariable == "EUtranCell":
				vis.ColorAxis.Expression = Eutrancellaxis
			else:
				vis.ColorAxis.Expression = Termpointsgw
def add_datatable(vis,vistitle,pagetilte,flagvariable,columnname,defaultcol,cellname):
	if Document.Properties["DisplayName"] in Eutrancellfdd:
		vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_RAW_DOD_Nodeview")                    
	elif Document.Properties["DisplayName"] in Eutrancellfddvector:					
		vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_V_RAW_DOD_Nodeview")                     
	elif Document.Properties["DisplayName"] in dodleveltables:
		vis.Data.DataTableReference = getdatatable("Multitable_DOD_Nodeview")   
	elif Document.Properties["DisplayName"] in fdduniontdd:
		vis.Data.DataTableReference = getdatatable(Fdduniontddtable)  
	elif Document.Properties["DisplayName"] in Termpoint:
		flagvariable = "TermPointToSGW"              
		vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_TERMPOINTTOSGW_V_RAW_DOD_Nodeview")  
	elif Document.Properties["DisplayName"] in Meankpis:
		vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_V_RAW_Mean")  
	elif Document.Properties["DisplayName"] in Multitable_kpis:
		vis.Data.DataTableReference = getdatatable("MultitablecelllevelKPI_Nodeview")
	if pagetilte == KPIcell:
		addaxistochart(vistitle,pagetilte,vis,flagvariable)
		vis.XAxis.ZoomRange = vis.XAxis.ZoomRange.DefaultRange
	elif pagetilte == KPInode:
		vis.XAxis.ZoomRange = vis.XAxis.ZoomRange.DefaultRange
	elif pagetilte == "Filtered Data - Node view":
		if flagvariable == "EUtranCell":
			addcolumns(vis,columnname,defaultcol)
		else:
			addcolumns(vis,columnname,cellname)

for page in Document.Pages:
	pagetilte = page.Title
	if page.Title == KPInode:
		for vis in page.Visuals:
			vistitle = vis.Title
			if "Selected" in vis.Title:
				vis = vis.As[Visualization]()
				add_datatable(vis,vistitle,pagetilte,Flag_variable,columnName,defaultcol,cellname)
			if upto_7days in vis.Title:
				vis = vis.As[Visualization]()
				add_datatable(vis,vistitle,pagetilte,Flag_variable,columnName,defaultcol,cellname)              
	elif page.Title == "KPI View - Cell":
		for vis in page.Visuals:
			vistitle = vis.Title
			if cellview  in vis.Title:
				vis = vis.As[Visualization]()
				add_datatable(vis,vistitle,pagetilte,Flag_variable,columnName,defaultcol,cellname)
			elif upto_7days in vis.Title:
				vis = vis.As[LineChart]()
				add_datatable(vis,vistitle,pagetilte,Flag_variable,columnName,defaultcol,cellname)
	elif page.Title == "Filtered Data - Node view":
		for vis in page.Visuals:
			vistitle = vis.Title
			if "Raw Filtered Data" in vis.Title:
				myVis = vis.As[Visualization]()
				myVis.TableColumns.Clear()
				add_datatable(myVis,vistitle,pagetilte,Flag_variable,columnName,defaultcol,cellname)
                