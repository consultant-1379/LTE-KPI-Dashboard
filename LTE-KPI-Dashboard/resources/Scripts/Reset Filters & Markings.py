from Spotfire.Dxp.Data import RowSelection,  IndexSet
from System.Collections.Generic import List,  Dictionary
from Spotfire.Dxp.Application.Visuals import Visualization, VisualTypeIdentifiers
from Spotfire.Dxp.Data import LimitingMarkingsEmptyBehavior
from Spotfire.Dxp.Application import Filters as filters



reset_pages_list = ['KPI View - Node', 'KPI View - Cell','Filtered Data - Node view']
        
def reset_marking_and_filtering(): 
	
    for data_table in Document.Data.Tables: 
        for marking in Document.Data.Markings: 
            rows = RowSelection(IndexSet(data_table.RowCount,  False))
            marking.SetSelection(rows,  data_table)
                
        for filter_scheme in Document.FilteringSchemes: 
            filter_scheme.ResetAllFilters()

def add_marking_holder(): 

    for reset_page_title in reset_pages_list:
        page = get_page(reset_page_title)
        for visualization in page.Visuals: 
            if visualization.TypeId == VisualTypeIdentifiers.BarChart: 
                visualization = visualization.As[Visualization]()
                visualization.Data.Filterings.Add(Document.Data.Markings["MarkingHolder"])
                if page.Title == 'KPI View - Cell':
					visualization.Data.LimitingMarkingsEmptyBehavior = LimitingMarkingsEmptyBehavior.ShowMessage
            elif visualization.TypeId == VisualTypeIdentifiers.Table:
                visualization = visualization.As[Visualization]()
                visualization.Data.Filterings.Add(Document.Data.Markings["MarkingHolder"])
                visualization.Data.LimitingMarkingsEmptyBehavior = LimitingMarkingsEmptyBehavior.ShowMessage
          
def get_page(page_name):

   for page in Document.Pages: 
      if page.Title == page_name: 
         return page
         
         
reset_marking_and_filtering()
add_marking_holder()