if($("#"+PrimaryElementDataUnit).text().includes("Kbps")){	
	var calculatedValue=$("#"+PrimaryElement);
	var unitElement=$("#"+PrimaryElementDataUnit);
	var deltaValue=$("#"+SecondaryElement);
	var threshold=999.999999;
	var bitToByteThreshold=1;
	var currentValue;
	var currentDeltaValue;
	var units = ["k", "M", "G", "T", "P", "E", "Z", "Y"];
	var bitByte="bps";
	var dataUnitToBps = $(".dataUnitToBps");

	calculatedValue.hide();
	unitElement.hide();
	deltaValue.hide();
	dataUnitToBps.hide();
	
	
	function Convert(magnitude){
		currentValue=(parseFloat(calculatedValue.text()));
					
					currentDeltaValue=(parseFloat(deltaValue.text()));
					
		currentValue=currentValue/bitToByteThreshold;
		while(currentValue>threshold){
			currentDeltaValue=currentDeltaValue/threshold;
			currentValue=currentValue/threshold;
			magnitude=magnitude+1;
		}

		calculatedValue.text(currentValue.toFixed([2]));
			if(String(currentDeltaValue).indexOf("NaN")!==-1){
			deltaValue.text("No Data");
		}else{
                if(currentDeltaValue>0){
					deltaValue.text("+"+currentDeltaValue.toFixed([2]));
                }else{	
                    deltaValue.text(currentDeltaValue.toFixed([2]));
                }
		}

		unitElement.text(" "+units[magnitude]+bitByte);
	}

	setTimeout(function(){
					if(calculatedValue.text().trim()!="No Data" && String(parseFloat(calculatedValue.text())).indexOf("NaN") ==-1){
						Convert(0);
					}
					calculatedValue.show();
					unitElement.show();
					deltaValue.show();
					dataUnitToBps.show();
	}, 2500);
}