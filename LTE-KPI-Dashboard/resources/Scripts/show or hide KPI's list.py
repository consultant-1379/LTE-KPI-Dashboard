import re
from Spotfire.Dxp.Application.Visuals import Visualization

for page in Document.Pages:
	if page.Title == "KPI Details":
		for vis in page.Visuals:
			if "Worst Performing Cells (Latest ROP)" in vis.Title:
				vis = vis.As[Visualization]()
				Rules = vis.TryGetFilterRules()
				if Document.Properties["DetailsDisplayName"] == "Average DL Packet Error Loss Rate"\
				or Document.Properties["DetailsDisplayName"] == "Average DL Packet Error Loss Rate per QCI"\
				or Document.Properties["DetailsDisplayName"] == "Average DL PDCP DRB LTE Latency per QCI"\
				or Document.Properties["DetailsDisplayName"] == "Average DL UE Latency"\
				or Document.Properties["DetailsDisplayName"] == "Average DL UE PDCP DRB Latency per QCI"\
				or Document.Properties["DetailsDisplayName"] == "Average UL Packet Loss Rate" \
				or Document.Properties["DetailsDisplayName"] == "Average UL Packet Loss Rate per QCI" \
				or Document.Properties["DetailsDisplayName"] == "GTP-U Downlink Packet Loss Rate" \
				or Document.Properties["DetailsDisplayName"] == "GTP-U Downlink Packet Out of Order Rate" \
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability (eNB) - Percentage Lost" \
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability (eNB) - Percentage Lost, the Second" \
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability - Percentage Lost"\
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability - Percentage Lost for Emergency Calls" \
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability - Percentage Lost per QCI" \
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability - Percentage Lost per QCI, eNodeB triggered only"\
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability - Session Time Normalized Loss Rate"\
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability - Session Time Normalized Loss Rate the Second" \
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability - Session Time Normalized per QCI Loss Rate" \
				or Document.Properties["DetailsDisplayName"] == "E-RAB Retainability - Session Time Normalized per QCI Loss Rate, the Second":				
					Rules[1][1].Enabled = True
					Rules[1][0].Enabled = False
					vis.XAxis.Reversed = True
				else:
					Rules[1][0].Enabled = True
					Rules[1][1].Enabled = False
					vis.XAxis.Reversed = False