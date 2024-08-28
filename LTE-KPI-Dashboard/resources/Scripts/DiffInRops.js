var Elements=new Array();

for (var i=1; i <= 12; i++) {
    if(i>=1&&i<=12){
        Elements[i-1]=[1,("#lastRopKPI"+i),("#lastWeekKPI"+i),("#diffIn2RopsKPI"+i)];
    }
}


function getDifference(lastRopVal, previousRopVal, differenceDiv){
	var rop1Val = $(lastRopVal).text();
	rop1Val = rop1Val.replace(/(\d)\D+$/g, '$1');
	var rop2Val = $(previousRopVal).text();
	var difference = rop1Val-rop2Val;
	
	if((Math.abs(parseFloat(rop1Val)-parseFloat(rop2Val))).toFixed(2)<0.00999999) {
        $(differenceDiv).css("color", "#d9dadb");
        $(differenceDiv).text(parseFloat(Math.round(difference * 100) / 100).toFixed(2));
    }else if(parseFloat(rop1Val)>parseFloat(rop2Val)){
        $(differenceDiv).css("color", "#a5c753");
        $(differenceDiv).text("+"+parseFloat(Math.round(difference * 100) / 100).toFixed(2));
    }else{
        $(differenceDiv).css("color", "#e95c38");
        $(differenceDiv).text(parseFloat(Math.round(difference * 100) / 100).toFixed(2));
    }

	
	if(isNaN(rop1Val) || isNaN(rop2Val)){
		$(differenceDiv).css("color", "#c8c8c8");
        $(differenceDiv).text("No Data");
	}
	
    $(differenceDiv).css({
        "font-size" : 16,
        "font-weight" : "bold"
    });
}
function getReversedDifference(lastRopVal, previousRopVal, differenceDiv){
	var rop1Val = $(lastRopVal).text();
	rop1Val = rop1Val.replace(/(\d)\D+$/g, '$1');
	var rop2Val = $(previousRopVal).text();
	var difference = rop1Val-rop2Val;
	
	if((Math.abs(parseFloat(rop1Val)-parseFloat(rop2Val))).toFixed(2)==0) {
        $(differenceDiv).css("color", "#d9dadb");
        $(differenceDiv).text(parseFloat(Math.round(difference * 100) / 100).toFixed(2));
    }else if(parseFloat(rop1Val)<parseFloat(rop2Val)){
        $(differenceDiv).css("color", "#a5c753");
        $(differenceDiv).text(parseFloat(Math.round(difference * 100) / 100).toFixed(2));
    }else{
        $(differenceDiv).css("color", "#e95c38");
        $(differenceDiv).text(parseFloat(Math.round(difference * 100) / 100).toFixed(2));
    }

	
	if(isNaN(rop1Val) || isNaN(rop2Val)){
		$(differenceDiv).css("color", "#c8c8c8");
        $(differenceDiv).text("No Data");
	}
	
    $(differenceDiv).css({
        "font-size" : 16,
        "font-weight" : "bold"
    });
}

setTimeout(function(){
		for (var i=1; i <= 12; i++) {		 
			 if( $("#KPIName"+i).text().includes("UL Packet Loss Captured in gNodeB")){
				 getReversedDifference(("#lastRopKPI"+i),("#lastWeekKPI"+i),("#diffIn2RopsKPI"+i))
				 continue;
			 }
				 getDifference(("#lastRopKPI"+i),("#lastWeekKPI"+i),("#diffIn2RopsKPI"+i));
		}
    }
, 1000);