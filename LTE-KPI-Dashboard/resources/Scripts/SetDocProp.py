from Spotfire.Dxp.Application.Visuals import Visualization
from Spotfire.Dxp.Application.Visuals import HtmlTextArea
from System import DateTime

kpi = Document.Properties["AccessibilityKPIs"]
    
names = ["${erabretainenbpercentlost}", "${erabretainenbpercentlostsec}", "${erabretainpercentlost}", "${erabretainpercentlostemergencycall}", "${erabretainpercentlostperqci}", \
"${erabretainpercentlostperqcienodeb}","${erabretainsestimenormlossrte}", "${erabretainsestimenormlossrtesec}","${erabretainsestimenormqcilossrte}", "${erabretainsestimenormqcilossrtesec}","${Inierabestabsuccrteemergcal}", \
"${Adederabestabsuccrte23ltr}","${Addederabestabsuccrteemergcallltr}","${Addederabestabsuccrteqci23ltr}","${Inierabestabsuccrte}","${Inierabestabsuccrtenomosignal}","${Inierabestabsuccrteqci19q4ltr}","${Ranacessmsg1succrte}", \
"${Robustrandaccsuccrte}","${cellhandovexesuccrte}","${cellhosuccrte}","${cellmobsuccrte}","${Hoexesuccrte}","${Hosuccrte}","${Intrafreqcellmobisuccrte}","${Mobisuccrte}","${Utransrvccsuccrte}",\
"${Avgdlpdcpcellthrough}","${Avgdlmaccellthrough}","${Avgdlpdcpdrbltelatencyqci}","${Avgdlpdcpuedrbthroghqci}","${Avgdlpdcpuethroughput}","${Avgdlpdcpuethroughcarrieragg}","${Avgdluelatency}","${Avgdluepdcpdrblatencyqci}",\
"${Avgmacueulthroughput}","${Avgulmaccellthroughput}","${Avgulpktlossrte}","${Avgulpktlossrteqci}","${Avgulpdcpcellthroughput}","${Avgulpdcpuethroughput}","${Avgulpdcpuethroughputcarrieragg}","${Avguluethroughputperlcg}","${GTPudownlinkpktlossrte}","${gtpudownlinkpktoutordrte}",\
"${Noravgulmaccellsuccpuschsubfrmeonlt}","${voipcellint}","${meanulpdcpuethroughput}","${meandlpdcpuethrough}","${cellavailabilty}","${Avgdlpkyerrlssrte}","${Avgdlpkterrlssrteqci}"]      

realNames = ["E-RAB Retainability (eNB) - Percentage Lost","E-RAB Retainability (eNB) - Percentage Lost, the Second", \
"E-RAB Retainability - Percentage Lost", "E-RAB Retainability - Percentage Lost for Emergency Calls", \
"E-RAB Retainability - Percentage Lost per QCI", "E-RAB Retainability - Percentage Lost per QCI, eNodeB triggered only", \
"E-RAB Retainability - Session Time Normalized Loss Rate", "E-RAB Retainability - Session Time Normalized Loss Rate the Second",\
 "E-RAB Retainability - Session Time Normalized per QCI Loss Rate", "E-RAB Retainability - Session Time Normalized per QCI Loss Rate, the Second",\
"Initial E-RAB Establishment Success Rate for Emergency Calls","Added E-RAB Establishment Success Rate 23.Q3 and later",\
"Added E-RAB Establishment Success Rate for Emergency Calls 23.Q3 and later","Added E-RAB Establishment Success Rate per QCI 23.Q3 and later","Initial E-RAB Establishment Success Rate",\
"Initial E-RAB Establishment Success Rate, No MO Signaling","Initial E-RAB Establishment Success Rate per QCI 19.Q4 and later","Random Access MSG1 Success Rate","Robust Random Access Success Rate",\
"Cell Handover Execution Success Rate","Cell Handover Success Rate","Cell Mobility Success Rate","Handover Execution Success Rate","Handover Success Rate","Intra-Frequency Cell Mobility Success Rate",\
"Mobility Success Rate","UTRAN SRVCC Success Rate","Average DL PDCP Cell Throughput","Average DL MAC Cell Throughput","Average DL PDCP DRB LTE Latency per QCI",\
"Average DL PDCP UE DRB Throughput per QCI","Average DL PDCP UE Throughput","Average DL PDCP UE Throughput for Carrier Aggregation","Average DL UE Latency",\
"Average DL UE PDCP DRB Latency per QCI","Average MAC UE UL THROUGHPUT","Average UL MAC Cell Throughput","Average UL Packet Loss Rate",\
"Average UL Packet Loss Rate per QCI","Average UL PDCP Cell Throughput","Average UL PDCP UE Throughput","Average UL PDCP UE Throughput for Carrier Aggregation",\
"Average UL UE Throughput per LCG","GTP-U Downlink Packet Loss Rate","GTP-U Downlink Packet Out of Order Rate","Normalized Average UL MAC Cell Throughput Considering Successful PUSCH subframe Only","VoIP Cell Integrity",\
"Mean UL PDCP UE Throughput","Mean DL PDCP UE Throughput","Cell Availability","Average DL Packet Error Loss Rate","Average DL Packet Error Loss Rate per QCI"]

for name in names:
	if kpi == name:
		Document.Properties["DisplayName"] = realNames[names.index(name)]               

if kpi == "${Avgdlpdcpcellthrough}" or kpi == "${Avgdlmaccellthrough}" or kpi == "${Avgdlpdcpuedrbthroghqci}" or kpi == "${Avgdlpdcpuethroughput}" or kpi == "${Avgdlpdcpuethroughcarrieragg}"\
or kpi == "${Avgdluelatency}" or kpi == "${Avgmacueulthroughput}" or kpi == "${Avgulmaccellthroughput}" or kpi == "${Avgulpdcpcellthroughput}" or kpi == "${Avgulpdcpuethroughput}" or kpi == "${Avgulpdcpuethroughputcarrieragg}"\
or kpi == "${Avguluethroughputperlcg}" or kpi == "${meanulpdcpuethroughput}" or kpi == "${meandlpdcpuethrough}":
	Document.Properties["Units"] = "kbps"
elif kpi == "${Ranacessmsg1succrte}" or kpi == "${Robustrandaccsuccrte}" :
	Document.Properties["Units"] = "UNDEFINED"
elif kpi == "${Noravgulmaccellsuccpuschsubfrmeonlt}" :
	Document.Properties["Units"] = "Mbps"
elif kpi == "${Avgdlpdcpdrbltelatencyqci}" or kpi == "${Avgdluepdcpdrblatencyqci}" or kpi == "${AvgoverDLlatency}":
	Document.Properties["Units"] = "Ms"
elif kpi == "${erabretainsestimenormlossrte}" or kpi == "${erabretainsestimenormlossrtesec}" or kpi == "${erabretainsestimenormqcilossrte}" or kpi == "${erabretainsestimenormqcilossrtesec}" :
	Document.Properties["Units"] = "1/s"
else:
	Document.Properties["Units"] = "%"
Document.Properties["triggerscript1"] = DateTime.Now