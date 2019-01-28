// JavaScript for multiple page update for ASPMaker 5+
// (C) 2006 e.World Technology Ltd.

var EW_AfterInit = false;

function EW_InitMultiPage(f) {
	EW_MaxPageIndex = 0;
	for (var i=0; i<arFldPage.length; i++) {
		if (arFldPage[i][1] > EW_MaxPageIndex)
			EW_MaxPageIndex = arFldPage[i][1]; 
	}	
	EW_MinPageIndex = EW_MaxPageIndex;
	for (var i=0; i<arFldPage.length; i++) {
		if (arFldPage[i][1] < EW_MinPageIndex)
			EW_MinPageIndex = arFldPage[i][1]; 
	}
	EW_NextPage(f);
	EW_AfterInit = true;
}

function EW_NextPage(f) {
	if (!(document.getElementById || document.all))
		return;
	if (EW_AfterInit && !EW_checkMyForm(f))
		return;
	EW_DisableButtons();
	var rowcnt = 0;
	while (rowcnt == 0 && EW_PageIndex < EW_MaxPageIndex) {
		EW_PageIndex++;
		var rowcnt;
		for (var i=0; i<arFldPage.length; i++)
			if (arFldPage[i][1] == EW_PageIndex) rowcnt++;
		if (rowcnt > 0) EW_ShowPage();
	}
	EW_UpdateButtons();
}

function EW_PrevPage() {
	if (!(document.getElementById || document.all))
		return;
	EW_DisableButtons();	
	var rowcnt = 0;
	while (rowcnt == 0 && EW_PageIndex > EW_MinPageIndex) {
		EW_PageIndex--;
		var rowcnt;
		for (var i=0; i<arFldPage.length; i++)
			if (arFldPage[i][1] == EW_PageIndex) rowcnt++;
		if (rowcnt > 0) EW_ShowPage();	
	}
	EW_UpdateButtons();
}

function EW_ShowPage() {
	for (var i=0; i<arFldPage.length; i++) {
		var row = EW_GetElement(arFldPage[i][0]);
		if (row) {
			row.style.display = (arFldPage[i][1] == EW_PageIndex) ? '' : 'none';
			if (row.style.display == '')
				EW_createEditor(arFldPage[i][0]);
		}	
	}
}

function EW_UpdateButtons() {
    var btn = EW_GetElement('btnPrevPage'); 
    if (btn) btn.disabled = EW_PageIndex <= EW_MinPageIndex;
    var btn = EW_GetElement('btnNextPage'); 
    if (btn) btn.disabled = EW_PageIndex >= EW_MaxPageIndex;
    var btn = EW_GetElement('btnAction'); 
    if (btn) btn.style.display = (EW_PageIndex < EW_MaxPageIndex) ? 'none' : ''; 
    var elem = EW_GetElement('ewPageInfo');
    if (elem) elem.innerHTML = EW_MultiPagePage + " " + (EW_PageIndex) + " " + EW_MultiPageOf + " " + (EW_MaxPageIndex);
}

function EW_DisableButtons() {
    var btn = EW_GetElement('btnPrevPage'); 
    if (btn) btn.disabled = false;
    var btn = EW_GetElement('btnNextPage'); 
    if (btn) btn.disabled = false;
    var btn = EW_GetElement('btnAction'); 
    if (btn) btn.style.display = 'none';    
}

function EW_GetElement(elemid) {
	return (document.all) ? document.all(elemid) :
		(document.getElementById) ? document.getElementById(elemid) : null;
}

function EW_IsElementVisible(elemid) {
	if (!(document.getElementById || document.all))
		return true;
	var elem = EW_GetElement(elemid);
	return (elem && elem.style.display == '');
}
