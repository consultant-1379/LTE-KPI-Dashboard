from System.Collections.Generic import List, Dictionary
from Spotfire.Dxp.Data import RowSelection, IndexSet

if len(Document.Properties["KPIList"].split(';')) == 12:
	Document.Properties["SelectionWarning"] = "Your 12 KPIs have been selected. To change your selection, click the Clear button"
else:
	Document.Properties["SelectionWarning"] = ""

dict = {
	"${erabretainenbpercentlost}":"E-RAB Retainability (eNB) - Percentage Lost",
	"${erabretainenbpercentlostsec}":"E-RAB Retainability (eNB) - Percentage Lost, the Second",
	"${erabretainpercentlost}":"E-RAB Retainability - Percentage Lost",
	"${erabretainpercentlostemergencycall}":"E-RAB Retainability - Percentage Lost for Emergency Calls",
	"${erabretainpercentlostperqci}":"E-RAB Retainability - Percentage Lost per QCI",  
	"${erabretainpercentlostperqcienodeb}":"E-RAB Retainability - Percentage Lost per QCI, eNodeB triggered only",
	"${erabretainsestimenormlossrte}":"E-RAB Retainability - Session Time Normalized Loss Rate",    
	"${erabretainsestimenormlossrtesec}":"E-RAB Retainability - Session Time Normalized Loss Rate the Second",
	"${erabretainsestimenormqcilossrte}":"E-RAB Retainability - Session Time Normalized per QCI Loss Rate",
	"${erabretainsestimenormqcilossrtesec}":"E-RAB Retainability - Session Time Normalized per QCI Loss Rate, the Second",
	"${Adederabestabsuccrte23ltr}":"Added E-RAB Establishment Success Rate 23.Q3 and later",
	"${Addederabestabsuccrteemergcallltr}":"Added E-RAB Establishment Success Rate for Emergency Calls 23.Q3 and later" , 
	"${Addederabestabsuccrteqci23ltr}":"Added E-RAB Establishment Success Rate per QCI 23.Q3 and later" ,
	"${Inierabestabsuccrte}":"Initial E-RAB Establishment Success Rate" ,
	"${Inierabestabsuccrteemergcal}":"Initial E-RAB Establishment Success Rate for Emergency Calls" ,
	"${Inierabestabsuccrtenomosignal}":"Initial E-RAB Establishment Success Rate, No MO Signaling",
	"${Inierabestabsuccrteqci19q4ltr}":"Initial E-RAB Establishment Success Rate per QCI 19.Q4 and later" ,
	"${Ranacessmsg1succrte}":"Random Access MSG1 Success Rate" ,
	"${Robustrandaccsuccrte}":"Robust Random Access Success Rate" ,
	"${Avgdlpdcpcellthrough}":"Average DL PDCP Cell Throughput" ,
	"${Avgdlmaccellthrough}":"Average DL MAC Cell Throughput",
	"${Avgdlpdcpdrbltelatencyqci}":"Average DL PDCP DRB LTE Latency per QCI", 
	"${Avgdlpdcpuedrbthroghqci}":"Average DL PDCP UE DRB Throughput per QCI",
	"${Avgdlpdcpuethroughput}":"Average DL PDCP UE Throughput",
	"${Avgdlpdcpuethroughcarrieragg}":"Average DL PDCP UE Throughput for Carrier Aggregation",
	"${Avgdluelatency}":"Average DL UE Latency",
	"${Avgdluepdcpdrblatencyqci}":"Average DL UE PDCP DRB Latency per QCI",
	"${Avgmacueulthroughput}":"Average MAC UE UL THROUGHPUT",
	"${Avgulmaccellthroughput}":"Average UL MAC Cell Throughput",
	"${Avgulpktlossrte}":"Average UL Packet Loss Rate" ,
	"${Avgulpktlossrteqci}":"Average UL Packet Loss Rate per QCI" ,
	"${Avgulpdcpcellthroughput}":"Average UL PDCP Cell Throughput" ,  
	"${Avgulpdcpuethroughput}":"Average UL PDCP UE Throughput" ,
	"${Avgulpdcpuethroughputcarrieragg}":"Average UL PDCP UE Throughput for Carrier Aggregation",
	"${Avguluethroughputperlcg}":"Average UL UE Throughput per LCG" , 
	"${GTPudownlinkpktlossrte}":"GTP-U Downlink Packet Loss Rate", 
	"${gtpudownlinkpktoutordrte}":"GTP-U Downlink Packet Out of Order Rate",
	"${Noravgulmaccellsuccpuschsubfrmeonlt}":"Normalized Average UL MAC Cell Throughput Considering Successful PUSCH subframe Only",
	"${voipcellint}":"VoIP Cell Integrity",
	"${cellhandovexesuccrte}":"Cell Handover Execution Success Rate",
	"${cellhosuccrte}":"Cell Handover Success Rate",   
	"${cellmobsuccrte}":"Cell Mobility Success Rate" ,
	"${Hoexesuccrte}":"Handover Execution Success Rate" ,
	"${Hosuccrte}":"Handover Success Rate"   ,
	"${Intrafreqcellmobisuccrte}":"Intra-Frequency Cell Mobility Success Rate" ,
	"${Mobisuccrte}":"Mobility Success Rate" ,
	"${Utransrvccsuccrte}":"UTRAN SRVCC Success Rate",
	"${cellavailabilty}":"Cell Availability",
	"${meanulpdcpuethroughput}":"Mean UL PDCP UE Throughput", 
	"${meandlpdcpuethrough}":"Mean DL PDCP UE Throughput"   
             
}

unitsDict = {
	"${erabretainenbpercentlost}":"%",
	"${erabretainenbpercentlostsec}":"%",
	"${erabretainpercentlost}":"%",
	"${erabretainpercentlostemergencycall}":"%",
    "${erabretainpercentlostperqci}":"%",
	"${erabretainpercentlostperqcienodeb}":"%",
	"${erabretainsestimenormlossrte}":"1/s",
	"${erabretainsestimenormlossrtesec}":"1/s",
    "${erabretainsestimenormqcilossrte}":"1/s",
	"${erabretainsestimenormqcilossrtesec}":"1/s",
	"${Adederabestabsuccrte23ltr}":"%",
	"${Addederabestabsuccrteemergcallltr}":"%" ,
	"${Addederabestabsuccrteqci23ltr}":"%" ,
	"${Inierabestabsuccrte}":"%" ,
	"${Inierabestabsuccrteemergcal}":"%" ,
	"${Inierabestabsuccrtenomosignal}":"%" ,
	"${Inierabestabsuccrteqci19q4ltr}":"%" ,
	"${Ranacessmsg1succrte}":"%" ,
	"${Robustrandaccsuccrte}":"%" ,
	"${Avgdlpdcpcellthrough}":"Kbps" ,
	"${Avgdlmaccellthrough}":"Kbps" ,
	"${Avgdlpdcpdrbltelatencyqci}":"Ms",
	"${Avgdlpdcpuedrbthroghqci}":"Kbps",
	"${Avgdlpdcpuethroughput}":"Kbps",
	"${Avgdlpdcpuethroughcarrieragg}":"Kbps",
	"${Avgdluelatency}":"Kbps",
	"${Avgdluepdcpdrblatencyqci}":"Ms",
	"${Avgmacueulthroughput}":"Kbps" ,
	"${Avgulmaccellthroughput}":"Kbps" ,
	"${Avgulpktlossrte}":"%" ,
	"${Avgulpktlossrteqci}":"%",
	"${Avgulpdcpcellthroughput}":"Kbps" , 
	"${Avgulpdcpuethroughput}":"Kbps" ,
	"${Avgulpdcpuethroughputcarrieragg}":"Kbps" , 
	"${Avguluethroughputperlcg}":"Kbps",
	"${GTPudownlinkpktlossrte}":"%" ,
	"${gtpudownlinkpktoutordrte}":"%" ,
	"${Noravgulmaccellsuccpuschsubfrmeonlt}":"Mbps" ,
	"${voipcellint}":"%" ,
	"${cellhandovexesuccrte}":"%" ,
	"${cellhosuccrte}":"%" ,
	"${cellmobsuccrte}":"%",
	"${Hoexesuccrte}":"%"  ,
	"${Hosuccrte}":"%" ,
	"${Intrafreqcellmobisuccrte}":"%" ,
	"${Mobisuccrte}":"%",
	"${Utransrvccsuccrte}":"%",
	"${cellavailabilty}":"%",
	"${meanulpdcpuethroughput}":"Kbps", 
	"${meandlpdcpuethrough}":"Kbps"    
  
}

counter = 1

for item in Document.Properties["KPISelection"]:
		if item not in Document.Properties["KPIList"] and len(Document.Properties["KPIList"].split(';')) < 12:
			if Document.Properties["KPIList"] != '':
				Document.Properties["KPIList"] += ';' + item
			else:
				Document.Properties["KPIList"] = item



for item in Document.Properties["KPIList"].split(';'):
	if len(Document.Properties["KPIList"].split(';')) <= 12:
		Document.Properties["SelectedKPI"+str(counter)] = item
		Document.Properties["SelectedKPIName"+str(counter)] = dict[item]
		Document.Properties["Units"+str(counter)] = unitsDict[item]
		if Document.Properties["SelectedKPIName"+str(counter)] == 'Mean UL PDCP UE Throughput' or Document.Properties["SelectedKPIName"+str(counter)] == 'Mean DL PDCP UE Throughput':
			Document.Properties["Note"+str(counter)] = "Note: This KPI is calculated based on day level information.\n The top value is the previous day's result and the bottom value is the difference between the previous day and the same day for the previous week."
		else:
			Document.Properties["Note"+str(counter)] = ""
		counter+=1

