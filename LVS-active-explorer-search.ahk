
/*
Adapted from [this post by fragman](<https://autohotkey.com/board/topic/19039-explorer-windows-manipulations/page-5#entry299720>).

If script is called with no arguments, then it's persistent, will sit waiting for a ^f in an explorer window to load the gui.
If script is called with any argument that evaluates to true, will try to search current active window, then close itself. This is
useful when calling it from another script

This uses COM to list the items in a folder, and it can be a little slow for folders with many files.
*/

#noenv

	groupadd, explorer, ahk_class ExploreWClass
	groupadd, explorer, ahk_class CabinetWClass

	LVS_Init("callback", "Filename")

	if %1%
	{
		ispersistent := False
		gosub Start
	}
	else
	{
		ispersistent := True
		hotkey, ifwinactive, ahk_group explorer
			hotkey, ^f, Start
		hotkey, ifwinactive
	}
return


Start:
	doc := getActiveExplorerWinDocument()
	if not doc
	{
		msgbox % "active window is not an explorer window. Aborting"
		gosub, fin
	}
	else
	{
		filelist := getFileList(doc)  ; this can take some time

		LV_Delete()
		LVS_SetBottomText(doc.Folder.Title)
		LVS_Show()  ; showing before populating because populating can take some time, can filter results before finishing
		LVS_SetList(filelist, "|")  ; filenames can't have |, can they? Populates LV with all files
		LVS_UpdateColOptions()  ; even with only one column, this will adjust width if there's a scrollbar
	}
return



callback(filelist, escaped := False) {
	global doc
	
	if not (escaped or filelist = "")
	{
		selectFiles(doc, filelist)
	}

	gosub, fin
}


getActiveExplorerWinDocument() {
	hWnd := winexist("a")
	
	if not winexist("ahk_id " hWnd " ahk_group explorer")
		return 0
	
	sa := ComObjCreate("Shell.Application")
	
	wins := sa.Windows
	loop % wins.count
	{
		window := wins.Item(A_Index-1)
		if(window.Hwnd = hWnd)
			break
	}
	doc := window.Document
	
	return doc
}


getFileList(doc) {
	items := doc.Folder.Items
	
	filelist := ""
	loop, % items.Count
	{
		item := items.Item(A_Index-1)
		itempath := item.Path
		splitpath, itempath, filename
		filelist .= filename "`n"
	}
	
	stringtrimright, filelist, filelist, 1  ; remove last `n
	return filelist
}


selectFiles(doc, filelist) {
	loop, parse, filelist, `n
	{
		item := doc.Folder.ParseName(A_LoopField)
		doc.SelectItem(item, A_Index = 1 ? 29 : 1) ; http://msdn.microsoft.com/en-us/library/bb774047(VS.85).aspx
	}
}


fin:
	if ispersistent
		LVS_Hide()
	else
		exitapp
return

#include %A_ScriptDir%\LVS.ahk