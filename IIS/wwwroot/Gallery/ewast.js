//
// Auto Suggest TextBox for ASPMaker 5+
// (C) 2006 e.World Technology Ltd.
//

// global variables
var g_nSelListItem = 0;
var g_sTextBoxID;
var g_bCancelSubmit;
var g_sOldTextBoxValue = "";
var g_MaxNewValueLength = 5; // only get data if value length <= this setting

function EW_astSetSelectedValue(sValue) {
	var hdnSelectedValue = document.getElementById("sv_" + g_sTextBoxID);
	hdnSelectedValue.value = sValue;
}

function EW_astSetTextBoxValue()	{
	var divListItem;
	divListItem = EW_astGetSelListItemDiv();
	if (divListItem)	{
		var sListItemValueID = GetDivListItemID(g_nSelListItem) + "_value";
		var hdnListItemValue = document.getElementById(sListItemValueID);		
		if (hdnListItemValue)
			EW_astSetSelectedValue(hdnListItemValue.value);		
		var txtCtrl = document.getElementById(g_sTextBoxID);
		txtCtrl.value = divListItem.innerHTML;
	}
}

function EW_astGetTextBoxValue()	{
	var txtCtrl = document.getElementById(g_sTextBoxID);
	return (txtCtrl) ? txtCtrl.value : '';
}	
		
function EW_astOnMouseClick(nListIndex, sTextBoxID, sDivID) {
	g_nSelListItem = nListIndex;
	g_sTextBoxID = sTextBoxID;					
	EW_astSetTextBoxValue();
	EW_astHideDiv(sDivID);
}

function EW_astOnMouseOver(nListIndex, sTextBoxID)	{
	g_sTextBoxID = sTextBoxID;				
	EW_astSelectListItem(nListIndex);
}	
	
function EW_astOnKeyPress(evt) {	
	if ((EW_astGetKey(evt) == 13) && (g_bCancelSubmit)) return false;		
	return true;
}	

function EW_astOnKeyUp(sTextBoxID, sDiv, evt) {	
	g_sTextBoxID = sTextBoxID;	
	var nKey = EW_astGetKey(evt);	
	// skip up/down/enter
	if ((nKey != 38) && (nKey != 40) && (nKey != 13))	{
		var sNewValue;
		sNewValue = EW_astGetTextBoxValue();			
		if ((sNewValue.length <= g_MaxNewValueLength) && (sNewValue.length > 0)) {
			if (nKey != 27) // skip escape									
				EW_ajaxupdatetextbox(sTextBoxID);					
		}			
		if (g_sOldTextBoxValue != sNewValue)
			EW_astSetSelectedValue("");					
	}
}	

function EW_astOnKeyDown(sTextBoxID, sDiv, evt) {	
	g_sTextBoxID = sTextBoxID;	
	// save current text box value before key press takes affect
	g_sOldTextBoxValue = EW_astGetTextBoxValue();	
	var nKey = EW_astGetKey(evt);					
	// handle up/down/enter
	if (nKey == 38) // up arrow		
		EW_astMoveDown();		
	else if (nKey == 40) // down arrow		
		EW_astMoveUp();
	else if (nKey == 13) { // enter
		// Note: Netscape will submit the form on pressing enter before firing the
		// keydown event. This only works with IE and FF.			
		if (EW_astIsVisibleDiv(sDiv)) {
			EW_astHideDiv(sDiv);						
			evt.cancelBubble = true;				
			if (evt.returnValue) evt.returnValue = false;
			if (evt.stopPropagation) evt.stopPropagation();			
			g_bCancelSubmit = true;
 		} else {
 			g_bCancelSubmit = false;
 		}
	} else {
		EW_astHideDiv(sDiv);
	}			
	return true;
}

function EW_astGetSelListItemDiv() {
	return EW_astGetListItemDiv(g_nSelListItem);
}			
		
function GetDivListItemID(nListItem) {
	return (g_sTextBoxID + "_mi_" + nListItem);
}

function EW_astGetListItemDiv(nListItem)	{
	var sDivListItemID;
	sDivListItemID = GetDivListItemID(nListItem);				
	return document.getElementById(sDivListItemID);
}		

function EW_astMoveUp() {
	var nListItem;
	nListItem = g_nSelListItem + 1;		
	if (EW_astGetListItemDiv(nListItem))	EW_astSelectListItem(nListItem);
}

function EW_astMoveDown() {
	var nListItem;
	nListItem = g_nSelListItem - 1;		
	if (nListItem != 0)	EW_astSelectListItem(nListItem);
}

function EW_astSelectListItem(nListItem)	{
	var divListItem;
	divListItem = EW_astGetListItemDiv(nListItem);					
	if(divListItem)	{
		if (nListItem != g_nSelListItem) {
			EW_astUnhighlightSelListItem();				
			g_nSelListItem = nListItem;
			EW_astSetTextBoxValue();						
			divListItem.className = "ewAstSelListItem";
		}
	}
}

function EW_astUnhighlightSelListItem() {
	var divListItem;
	divListItem = EW_astGetSelListItemDiv();	
	if (divListItem) divListItem.className = "ewAstListItem";		
}

function EW_astGetKey(evt) {
	evt = (evt) ? evt : (window.event) ? event : null;
	if (evt) {
		var cCode = (evt.charCode) ? evt.charCode :
				((evt.keyCode) ? evt.keyCode :
				((evt.which) ? evt.which : 0));
		return cCode; 
	}
}

function EW_astHideDiv(sDivID) {	
	document.getElementById(sDivID).style.visibility = 'hidden';
	g_nSelListItem = 0;
}

function EW_astIsVisibleDiv(sDivID) {
	return document.getElementById(sDivID).style.visibility != 'hidden';		
}

function EW_astShowDiv(sDivID, sDivContent) {	
	var divList;
	divList = document.getElementById(sDivID);		
	var sInnerHtml;
	// use iframe with the same size as the div		
	if (document.all) { // IE
		sInnerHtml = "<div id='" + sDivID + "_content' style='z-index:999; position:absolute;'>";
		sInnerHtml += sDivContent;
		sInnerHtml += "</div><iframe id='" + sDivID + "_iframe' src='about:blank' frameborder='1' scrolling='no'></iframe>";
		divList.innerHTML = sInnerHtml;
		var divContent = document.getElementById(sDivID + "_content");			
		var divIframe = document.getElementById(sDivID + "_iframe");					
		divContent.className = "ewAstList";
		divList.className = "ewAstListBase";				
		divIframe.style.width = divContent.offsetWidth + 'px';
		divIframe.style.height = divContent.offsetHeight + 'px';
		divIframe.marginTop = "-" + divContent.offsetHeight + 'px';
	} else {
		divList.innerHTML = sDivContent;	
	}	
	divList.style.visibility = 'visible';		
}
