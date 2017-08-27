
Requires [LVS.ahk](https://github.com/diogo0258/lvs)

Adapted from [this post by fragman](<https://autohotkey.com/board/topic/19039-explorer-windows-manipulations/page-5#entry299720>).

Lists files from active Explorer window in a searchable listview.

Pressing Enter with filenames selected in listview will select those files in original Explorer window.

Useful to select multiple files via regex, for example, in order to copy/move/delete them.

```
; If script is called with no arguments, then it's persistent, will sit waiting for a ^f in a explorer window to load the gui.
; If script is called with any argument that evaluates to true, will try to search current active window, then close itself. This is
; useful when calling it from another script
```

This uses COM to list the items in a folder, and it can be a little slow for folders with many files.

![demo](https://github.com/diogo0258/lvs-active-explorer-search/raw/master/demo/demo-gif.gif)
