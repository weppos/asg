// Function to open a popup window
function openWin(winURL,winName,winFeatures) {
  	window.open(winURL,winName,winFeatures);
}

// Function to swap a layer visibility status
function swapDisplayLayer(layername) {
	var blnDisplay = document.getElementById(layername).style.display;
	if (blnDisplay == 'none') { document.getElementById(layername).style.display = 'block'; }
	else { document.getElementById(layername).style.display = 'none'; }
}

// Function to swap a tr visibility status
function swapDisplayTr(layername) {
	var blnDisplay = document.getElementById(layername).style.display;
	if (blnDisplay == 'none') { document.getElementById(layername).style.display = 'table-row'; }
	else { document.getElementById(layername).style.display = 'none'; }
}