from Spotfire.Dxp.Framework.ApplicationModel import NotificationService
from Spotfire.Dxp.Framework.ApplicationModel import ProgressService, ProgressCanceledException
from Spotfire.Dxp.Application import Filters as filters
from Spotfire.Dxp.Application.Visuals import Visualization,VisualTypeIdentifiers
from Spotfire.Dxp.Data.Import import DataTableDataSource
from Spotfire.Dxp.Data import RowSelection, IndexSet
from Spotfire.Dxp.Data import AddRowsSettings
from datetime import datetime
from Spotfire.Dxp.Data import LimitingMarkingsEmptyBehavior
import time
from System import DateTime

ps = Application.GetService[ProgressService]()
notify = Application.GetService[NotificationService]() 
reset_pages_list = ['KPI View - Node', 'KPI View - Cell','Filtered Data - Node view']

def refresh_data():
    ps.CurrentProgress.ExecuteSubtask('Starting to refresh data')
    if Document.Properties["DisplayName"] == "E-RAB Retainability (eNB) - Percentage Lost"\
    or Document.Properties["DisplayName"] == "E-RAB Retainability (eNB) - Percentage Lost, the Second" \
    or Document.Properties["DisplayName"] == "E-RAB Retainability - Percentage Lost" \
    or Document.Properties["DisplayName"] == "E-RAB Retainability - Session Time Normalized Loss Rate" \
    or Document.Properties["DisplayName"] == "E-RAB Retainability - Session Time Normalized Loss Rate the Second"\
    or Document.Properties["DisplayName"] == "Random Access MSG1 Success Rate"\
    or Document.Properties["DisplayName"] == "Added E-RAB Establishment Success Rate 23.Q3 and later"\
    or Document.Properties["DisplayName"] == "Initial E-RAB Establishment Success Rate"\
    or Document.Properties["DisplayName"] == "Initial E-RAB Establishment Success Rate, No MO Signaling"\
    or Document.Properties["DisplayName"] == "Average DL MAC Cell Throughput"\
    or Document.Properties["DisplayName"] == "Average DL PDCP Cell Throughput"\
    or Document.Properties["DisplayName"] == "Average DL PDCP UE Throughput"\
    or Document.Properties["DisplayName"] == "Average DL PDCP UE Throughput for Carrier Aggregation"\
    or Document.Properties["DisplayName"] == "Average DL UE Latency"\
    or Document.Properties["DisplayName"] == "Average MAC UE UL THROUGHPUT"\
    or Document.Properties["DisplayName"] == "Average UL MAC Cell Throughput"\
    or Document.Properties["DisplayName"] == "Average UL Packet Loss Rate"\
    or Document.Properties["DisplayName"] == "Average UL PDCP Cell Throughput"\
    or Document.Properties["DisplayName"] == "Average UL PDCP UE Throughput"\
    or Document.Properties["DisplayName"] == "Average UL PDCP UE Throughput for Carrier Aggregation"\
    or Document.Properties["DisplayName"] == "VoIP Cell Integrity"\
    or Document.Properties["DisplayName"] == "Cell Availability"\
    or Document.Properties["DisplayName"] == "Normalized Average UL MAC Cell Throughput Considering Successful PUSCH subframe Only"\
    or Document.Properties["DisplayName"] == "Robust Random Access Success Rate":
        Document.Data.Tables["IL_DC_E_ERBS_EUTRANCELLFDD_RAW_DOD_Nodeview"].ReloadAllData()
    elif Document.Properties["DisplayName"] == "E-RAB Retainability - Percentage Lost for Emergency Calls"\
    or Document.Properties["DisplayName"] == "E-RAB Retainability - Percentage Lost per QCI" \
    or Document.Properties["DisplayName"] == "E-RAB Retainability - Percentage Lost per QCI, eNodeB triggered only" \
    or Document.Properties["DisplayName"] == "E-RAB Retainability - Session Time Normalized per QCI Loss Rate" \
    or Document.Properties["DisplayName"] == "Added E-RAB Establishment Success Rate for Emergency Calls 23.Q3 and later" \
    or Document.Properties["DisplayName"] == "Added E-RAB Establishment Success Rate per QCI 23.Q3 and later" \
    or Document.Properties["DisplayName"] == "Average DL PDCP DRB LTE Latency per QCI"\
    or Document.Properties["DisplayName"] == "Average DL PDCP UE DRB Throughput per QCI"\
    or Document.Properties["DisplayName"] == "Average DL UE PDCP DRB Latency per QCI"\
    or Document.Properties["DisplayName"] == "Average UL Packet Loss Rate per QCI"\
    or Document.Properties["DisplayName"] == "Average UL UE Throughput per LCG"\
    or Document.Properties["DisplayName"] == "E-RAB Retainability - Session Time Normalized per QCI Loss Rate, the Second":	
        Document.Data.Tables["IL_DC_E_ERBS_EUTRANCELLFDD_V_RAW_DOD_Nodeview"].ReloadAllData()
    elif Document.Properties["DisplayName"] == "Initial E-RAB Establishment Success Rate for Emergency Calls"\
    or Document.Properties["DisplayName"] == "Initial E-RAB Establishment Success Rate per QCI 19.Q4 and later":
        Document.Data.Tables["IL_DC_E_ERBS_EUTRANCELLFDD_V_RAW_DOD_Nodeview"].ReloadAllData()
        Document.Data.Tables["IL_DC_E_ERBS_EUTRANCELLFDD_RAW_DOD_Nodeview"].ReloadAllData()
    elif Document.Properties["DisplayName"] == "Cell Handover Execution Success Rate"\
    or Document.Properties["DisplayName"] == "Cell Handover Success Rate" \
    or Document.Properties["DisplayName"] == "Cell Mobility Success Rate" \
    or Document.Properties["DisplayName"] == "Handover Execution Success Rate" \
    or Document.Properties["DisplayName"] == "Handover Success Rate" \
    or Document.Properties["DisplayName"] == "Intra-Frequency Cell Mobility Success Rate" \
    or Document.Properties["DisplayName"] == "Mobility Success Rate"\
    or Document.Properties["DisplayName"] == "UTRAN SRVCC Success Rate":  
        Document.Data.Tables["FDD_DOD_KPI's_Nodeview"].ReloadAllData()
        Document.Data.Tables["TDD_DOD_KPI's_Nodeview"].ReloadAllData()
    elif Document.Properties["DisplayName"] == "GTP-U Downlink Packet Loss Rate"\
    or Document.Properties["DisplayName"] == "GTP-U Downlink Packet Out of Order Rate":  
        Document.Data.Tables["IL_DC_E_ERBS_TERMPOINTTOSGW_V_RAW_DOD_Nodeview"].ReloadAllData()
    elif Document.Properties["DisplayName"] == "Mean UL PDCP UE Throughput"\
    or Document.Properties["DisplayName"] == "Mean DL PDCP UE Throughput": 
        Document.Data.Tables["IL_DC_E_ERBS_EUTRANCELLFDD_V_RAW_Mean"].ReloadAllData()
    elif Document.Properties["DisplayName"] == "Average DL Packet Error Loss Rate"\
    or Document.Properties["DisplayName"] == "Average DL Packet Error Loss Rate per QCI": 
        Document.Data.Tables["MultitablecelllevelKPI_Nodeview"].ReloadAllData()

def remove_filters_and_markings():

    for reset_page_title in reset_pages_list:
        page = get_page(reset_page_title)
        for visualization in page.Visuals:
            if visualization.TypeId == VisualTypeIdentifiers.BarChart:
                visualization = visualization.As[Visualization]()
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                if page.Title == 'KPI View - Cell':
					visualization.Data.LimitingMarkingsEmptyBehavior = LimitingMarkingsEmptyBehavior.ShowAll
            elif visualization.TypeId == VisualTypeIdentifiers.Table:
                visualization = visualization.As[Visualization]()
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                visualization.Data.Filterings.Remove(Document.Data.Markings["MarkingHolder"])
                visualization.Data.LimitingMarkingsEmptyBehavior = LimitingMarkingsEmptyBehavior.ShowAll

def get_page(page_name):
    """
    Returns page object based on received page name
    Arguments:
       page_name {String} -- page name string
    Returns:
        page {obj} -- page object
    """

    for page in Document.Pages: 
        if page.Title == page_name: 
            return page     
            
def main():
    """
    Entry Point of script, call to refresh data for DoD data tables
    """
    try:
        refresh_data()
        remove_filters_and_markings()
    except Exception as e:
        msg = "Something Went Wrong"
        notify.AddWarningNotification("Exception","Error in fetching data",msg)
        print("Exception: ", e)


ps.ExecuteWithProgress('Fetching data', 'Please be patient, this can take several minutes...', main)