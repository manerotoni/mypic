//Properties that can be set by the user
//property of overlayd points on image
ROIsize= 6; //a ROIsizexROIsize is averaged use even numbers
labelSize = 5; //little dot on overlay
setForegroundColor(255,0,0); //color of dot
colorText = "red"; //color of Text
format = "jpeg"; //format output overlay png (small-size, good look), tiff (bigsize, goodlook), jpeg (smallsize, badlook)
setBatchMode(true); // if true no images are shown and macro runs faster

//Start the macro

//ROI's where to measure intensity
ROIs = newArray(1);
ROIs[0] = ROIsize; //e.g. 12x12 ROI

//What to measure and to mark
run("Point Tool...", "mark=" + labelSize + " label selection=" + colorText);
run("Set Measurements...", "area mean standard centroid redirect=None decimal=6");

dir = getDirectory("Choose a Directory "); 

//a text file to write all the results
fileout = dir  + "fluorescenceMeasure" + ".txt";
if(File.exists(fileout)){
	File.delete(fileout);
}

File.append("Pt_or_Roi \t Mean \t Std \t X \t Y",fileout)

processFiles(dir, fileout, ROIs, colorText, format);

function computeFiles(filelsm, fileout, ROIs, color, format) {
	//clean up	
	run("Close All");
	roiManager("reset"); 
	run("Clear Results");
	
	open(filelsm);

	//this is the dimension in pixels!!
	getDimensions(width, height, channels, slices, frames); 
	getPixelSize(unit, pixelWidth, pixelHeight);
	idx = indexOf(File.nameWithoutExtension,"_preFCS");
	//this is the txt file containining the coordinates
	filetxt = File.directory + substring(File.nameWithoutExtension,0,idx) + ".txt";
	filejpg = File.directory +  File.nameWithoutExtension + ".jpg";
	filestring = File.openAsString(filetxt);
	rows = split(filestring,"\n");
	x=newArray(rows.length-1);
	y=newArray(rows.length-1); 
	z=newArray(rows.length-1); 
	
	for(i=0; i<rows.length-1; i++) {
		columns=split(rows[i+1]); //first line is comment
		x[i]=parseInt(columns[0]);
		y[i]=parseInt(columns[1]); 
	 	z[i]=parseInt(columns[2]);
	 	// 0x0 is center of image
	 	posX = x[i]/pixelWidth + width/2;
	 	posY = y[i]/pixelHeight + height/2;
	 	//point of FCS-recording
	 	makePoint(posX, posY);
	 	roiManager("Add");
	 	for(j=0; j<ROIs.length; j++) {
	 		makeRectangle(posX - (ROIs[j]+1)/2, posY-(ROIs[j]+1)/2, ROIs[j], ROIs[j]); 
			roiManager("Add");
		}
	}
	roiManager("Measure");
	File.append(filelsm, fileout);
	for(i=0;i<nResults/2;i=i+2) {
		File.append("FCSpt \t" + getResult("Mean",i) + "\t" + getResult("StdDev",i) + "\t " + getResult("X",i) + "\t" + getResult("Y",i), fileout);
		File.append("ROI_"+ ROIs[0] + "x" + ROIs[0] + "px" + "\t" + getResult("Mean",i+1) + "\t" + getResult("StdDev",i+1) + "\t " + getResult("X",i+1) + "\t " + getResult("Y",i+1), fileout);
	}
	run("From ROI Manager");
	saveAs("Tiff", File.directory +  File.nameWithoutExtension + ".tif");
	roiManager("reset");
	run("RGB Color");
	//First create a jpg of image with point where measurment was taken and ROI for 
	for(i=0; i < rows.length - 1; i++) {
	 	// 0x0 is center of image
	 	posX = x[i]/pixelWidth + width/2;
	 	posY = y[i]/pixelHeight + height/2;
	 	//point of FCS-recording
	 	makePoint(posX, posY);
	 	roiManager("Add");
	 	roiManager("Draw"); 
	 	setFont("SansSerifBold", 14,"Bold");
	 	setColor(color);
	 	drawString(i+1,posX+2,posY-2);
	}
	saveAs(format, filejpg);
	
} 

function processFiles(dir, fileout, ROIs, color, format) {
	list = getFileList(dir);
	for (i = 0; i < list.length; i++) {
		if(endsWith(list[i],"preFCS.lsm")) {
			computeFiles(list[i], fileout, ROIs, color, format);
		}
	}
}
