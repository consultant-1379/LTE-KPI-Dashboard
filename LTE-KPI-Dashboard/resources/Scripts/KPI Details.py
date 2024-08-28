from Spotfire.Dxp.Application.Visuals import Visualization,TablePlot, TablePlotColumnSortMode
import re
from System import DateTime

myVis = myVis.As[Visualization]()
dataTable = myVis.Data.DataTableReference
myVis.TableColumns.Clear()
myVis.SortInfos.Clear()
activeVisual = Document.ActivePageReference.ActiveVisualReference
Flag_variable = "EUtranCell"
Flag_variable1 = "Datetime"
Eutrancellfdd = ["E-RAB Retainability (eNB) - Percentage Lost","E-RAB Retainability (eNB) - Percentage Lost, the Second","E-RAB Retainability - Percentage Lost",\
"E-RAB Retainability - Session Time Normalized Loss Rate","E-RAB Retainability - Session Time Normalized Loss Rate the Second","Random Access MSG1 Success Rate",\
"Added E-RAB Establishment Success Rate 23.Q3 and later","Initial E-RAB Establishment Success Rate","Initial E-RAB Establishment Success Rate, No MO Signaling",\
"Average DL MAC Cell Throughput","Average DL PDCP Cell Throughput","Average DL PDCP UE Throughput","Average DL PDCP UE Throughput for Carrier Aggregation","Average DL UE Latency",\
"Average MAC UE UL THROUGHPUT","Average UL MAC Cell Throughput","Average UL Packet Loss Rate","Average UL PDCP Cell Throughput","Average UL PDCP UE Throughput",\
"Average UL PDCP UE Throughput for Carrier Aggregation","VoIP Cell Integrity","Cell Availability","Normalized Average UL MAC Cell Throughput Considering Successful PUSCH subframe Only",\
"Robust Random Access Success Rate"]
Eutrancellfddvector = ["E-RAB Retainability - Percentage Lost for Emergency Calls","E-RAB Retainability - Percentage Lost per QCI","E-RAB Retainability - Percentage Lost per QCI, eNodeB triggered only",\
"E-RAB Retainability - Session Time Normalized per QCI Loss Rate","Added E-RAB Establishment Success Rate for Emergency Calls 23.Q3 and later",\
"Added E-RAB Establishment Success Rate per QCI 23.Q3 and later","Average DL PDCP DRB LTE Latency per QCI","Average DL PDCP UE DRB Throughput per QCI",\
"Average DL UE PDCP DRB Latency per QCI","Average UL Packet Loss Rate per QCI","Average UL UE Throughput per LCG","E-RAB Retainability - Session Time Normalized per QCI Loss Rate, the Second"]
fdduniontdd = ["Cell Handover Execution Success Rate","Cell Handover Success Rate","Cell Mobility Success Rate","Handover Execution Success Rate","Handover Success Rate",\
"Intra-Frequency Cell Mobility Success Rate","Mobility Success Rate","UTRAN SRVCC Success Rate"]
dodleveltables = ["Initial E-RAB Establishment Success Rate for Emergency Calls","Initial E-RAB Establishment Success Rate per QCI 19.Q4 and later"]
Termpoint = ["GTP-U Downlink Packet Out of Order Rate","GTP-U Downlink Packet Loss Rate"]
MeanKpis = ["Mean UL PDCP UE Throughput","Mean DL PDCP UE Throughput"]
chartmarking = "KPIDetailsMarking"
date_time = "DATETIME_ID"
date_id = "DATE_ID"
Eutran = "<[EUtranCell]>"

def getdatatable(tablename):
	try:
		return Document.Data.Tables[tablename]
	except KeyError:
		raise ("Error - cannot find data table: " + tablename)

def remove_marking(vis,markingname):
	vis.Data.Filterings.Remove(Document.Data.Markings[markingname])
	vis.Data.Filterings.Remove(Document.Data.Markings[markingname])
	vis.Data.Filterings.Remove(Document.Data.Markings[markingname])
	vis.Data.Filterings.Remove(Document.Data.Markings[markingname])
	Document.Properties["aggregationforcellchart"] = "ROP"
	Document.Properties["aggregationfortable"] = "Raw"

def add_marking(vis,markingname):
	vis.Data.Filterings.Add(Document.Data.Markings[markingname])
	Document.Properties["aggregationforcellchart"] = "DAY"
	Document.Properties["aggregationfortable"] = "Day"

def add_datalimit(vis):

	if Document.Properties["DetailsDisplayName"] == "Mean UL PDCP UE Throughput":
		vis.Data.WhereClauseExpression = "[DCVECTOR_INDEX]<=26 and [KPI Value] is not null"
	else:
		vis.Data.WhereClauseExpression = "[KPI Value] is not null"
def add_sorting(datetime,myvis,datatable):
    myvis.SortInfos.Add(datatable.Columns[datetime],TablePlotColumnSortMode.Ascending)

def add_values_to_networkchart(vis,datetime):
	vis.XAxis.Expression = '<['+datetime+']>'
	vis.Data.WhereClauseExpression = '['+ datetime + '] is not null'
	vis.XAxis.ZoomRange = vis.XAxis.ZoomRange.DefaultRange
try:
    Document.Properties["DetailsKPI"] = Document.Properties["SelectedKPI"+str(activeVisual.Title)]
    Document.Properties["DetailsKPI2"] = Document.Properties["SelectedKPI"+str(activeVisual.Title)]
    Document.Properties["DetailsDisplayName"] = Document.Properties["SelectedKPIName"+str(activeVisual.Title)]
    Document.Properties["UnitsSelected"] = Document.Properties["Units"+str(activeVisual.Title)]
    kpiFormula = Document.Properties[Document.Properties["SelectedKPI"+str(activeVisual.Title)][2:-1]]
    columnName = re.findall(r'\[([^[]+)\]', kpiFormula)
    columnName = list(set(columnName))
    if Document.Properties["DetailsDisplayName"] in Eutrancellfdd:
        myVis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_RAW_DOD")
        dataTable = myVis.Data.DataTableReference
    elif Document.Properties["DetailsDisplayName"] in Eutrancellfddvector:
        myVis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_V_RAW_DOD")
        dataTable = myVis.Data.DataTableReference
    elif Document.Properties["DetailsDisplayName"] in fdduniontdd:
        myVis.Data.DataTableReference = getdatatable("FDD_Union_TDD_DOD_KPI's")
        dataTable = myVis.Data.DataTableReference
    elif Document.Properties["DetailsDisplayName"] in dodleveltables:
        myVis.Data.DataTableReference = getdatatable("DOD_LEVEL_Tables")
        dataTable = myVis.Data.DataTableReference
    elif Document.Properties["DetailsDisplayName"] in Termpoint:
        myVis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_TERMPOINTTOSGW_V_RAW_DOD")
        dataTable = myVis.Data.DataTableReference
        Flag_variable = "TermPointToSGW"
    else:
        Flag_variable1 = "Date"
        myVis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_V_DAY_NW")
        dataTable = myVis.Data.DataTableReference

    if Flag_variable == "EUtranCell" and Flag_variable1 == "Datetime":
        myVis.TableColumns.Add(dataTable.Columns.Item["DATETIME_ID"])
        myVis.TableColumns.Add(dataTable.Columns.Item["ERBS"])
        myVis.TableColumns.Add(dataTable.Columns.Item["EUtranCell"])
        remove_marking(myVis,chartmarking)
        add_datalimit(myVis)
        add_sorting(date_time,myVis,dataTable)
    elif Flag_variable == "EUtranCell" and Flag_variable1 == "Date":
        myVis.TableColumns.Add(dataTable.Columns.Item["DATE_ID"])
        myVis.TableColumns.Add(dataTable.Columns.Item["ERBS"])
        myVis.TableColumns.Add(dataTable.Columns.Item["EUtranCell"])
        myVis.TableColumns.Add(dataTable.Columns.Item["DCVECTOR_INDEX"])
        add_marking(myVis,chartmarking)
        add_datalimit(myVis)
        add_sorting(date_id,myVis,dataTable)
    else:
        myVis.TableColumns.Add(dataTable.Columns.Item["DATETIME_ID"])
        myVis.TableColumns.Add(dataTable.Columns.Item["ERBS"])
        myVis.TableColumns.Add(dataTable.Columns.Item["TermPointToSGW"])
        remove_marking(myVis,chartmarking)
        add_datalimit(myVis)
        add_sorting(date_time,myVis,dataTable)
    myVis.TableColumns.Add(dataTable.Columns.Item["KPI Value"])
    for item in columnName:
        if item != "DATETIME_ID":
            myVis.TableColumns.Add(dataTable.Columns.Item[item])
    Document.ActivePageReference = Document.Pages[2]
except KeyError:
	print ("Value Error - No KPI selected")
except ValueError:
	print ("Can't set KPI document property")


for page in Document.Pages:
	if page.Title == "KPI Details":
		for vis in page.Visuals:
			if "Worst Performing Cells" in vis.Title:
				vis = vis.As[Visualization]()
				Rules = vis.TryGetFilterRules()
				if Document.Properties["DetailsDisplayName"] in Eutrancellfdd:                
					vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELL(FDD&TDD)_RAW") 
				elif Document.Properties["DetailsDisplayName"] in Eutrancellfddvector:
					vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_V_RAW") 
				elif Document.Properties["DetailsDisplayName"] in fdduniontdd:
					vis.Data.DataTableReference = getdatatable("Mobility_FDD_Union_TDD_RAW")  
				elif Document.Properties["DetailsDisplayName"] in dodleveltables:
					vis.Data.DataTableReference = getdatatable("RAW_LEVEL_Tables") 
				elif Document.Properties["DetailsDisplayName"] in Termpoint:
					vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_TERMPOINTTOSGW_V_RAW")
				else:
					vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_V_DAY")
				if Flag_variable == "EUtranCell":
					vis.XAxis.Expression = "<[EUtranCell] NEST [ERBS]>"
					vis.ColorAxis.Expression = Eutran
				else:
					vis.XAxis.Expression = "<[TermPointToSGW] NEST [ERBS]>"
					vis.ColorAxis.Expression = "<[TermPointToSGW]>"
			elif "Up to 7 Days" in vis.Title:
				vis = vis.As[Visualization]()
				if Document.Properties["DetailsDisplayName"] in Eutrancellfdd: 
					vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_RAW_DOD")
				elif Document.Properties["DetailsDisplayName"] in Eutrancellfddvector:
					vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_V_RAW_DOD")
				elif Document.Properties["DetailsDisplayName"] in fdduniontdd:
					vis.Data.DataTableReference = getdatatable("FDD_Union_TDD_DOD_KPI's") 
				elif Document.Properties["DetailsDisplayName"] in dodleveltables:
					vis.Data.DataTableReference = getdatatable("DOD_LEVEL_Tables")
				elif Document.Properties["DetailsDisplayName"] in Termpoint:
					vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_TERMPOINTTOSGW_V_RAW_DOD") 
				else:
					vis.Data.DataTableReference = getdatatable("IL_DC_E_ERBS_EUTRANCELLFDD_V_DAY_NW")
				if Flag_variable == "EUtranCell" and Flag_variable1 == "Datetime":
					vis.ColorAxis.Expression = Eutran
					vis.XAxis.Expression = "<[DATETIME_ID]>"
					remove_marking(vis,chartmarking)
				elif Flag_variable == "EUtranCell" and Flag_variable1 == "Date":
					vis.ColorAxis.Expression = Eutran
					vis.XAxis.Expression = "<[DATE_ID]>"
					add_marking(vis,chartmarking)
				else:
					vis.ColorAxis.Expression = "<[TermPointToSGW]>"
					vis.XAxis.Expression = "<[DATETIME_ID]>"
					remove_marking(vis,chartmarking)
				vis.XAxis.ZoomRange = vis.XAxis.ZoomRange.DefaultRange
			elif "Network View (Up to  7 Days)" in vis.Title:
				vis = vis.As[Visualization]()
				if Flag_variable1 == "Datetime":
					add_values_to_networkchart(vis,date_time)
				elif Flag_variable1 == "Date":
					add_values_to_networkchart(vis,date_id)
Document.Properties["triggerscript"] = DateTime.Now
