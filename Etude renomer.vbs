' Anciens nom d'étude
Anciens_Nom = "VIL_PM"

' Nouveaux nom d'étude
Nouveaux_Nom = "MTT_PM"

' Chemin d'étude pour renommer
DossierEtudePA = "D:\PROJETS\QGIS\2020-11-05 PA-64203-000A_Docs\Comac"



Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = False 

ShowSubfolders objFSO.GetFolder(DossierEtudePA)

objExcel.Quit
Set objExcel = Nothing

WScript.Echo "Terminer"

Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
		sNewSubFolder = Subfolder.Name
		sNewSubFolder = Replace(sNewSubFolder, Anciens_Nom, Nouveaux_Nom)
		if (sNewSubFolder<>Subfolder.Name) then 
			Subfolder.Move(Subfolder.ParentFolder+"\"+sNewSubFolder)
		end if
			
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
			sNewFile = objFile.Name
			sNewFile = Replace(sNewFile, Anciens_Nom, Nouveaux_Nom)
			
			If InStr(objFile.Name, "-Plan") <> 0 Then
				Set objWorkbook = objExcel.Workbooks.Open(objFile.ParentFolder + "\" + objFile.Name)
				objExcel.Worksheets(1).Range("C8").Value = sNewFile
				objExcel.Worksheets(2).Range("B1").Value = sNewFile
				objExcel.Worksheets(3).Range("D1").Value = sNewFile
				objExcel.Worksheets(4).Range("B6").Value = sNewFile
				objWorkbook.Save
				objWorkbook.Close 
			End If
			
			if (sNewFile<>objFile.Name) then 
				objFile.Move(objFile.ParentFolder + "\" + sNewFile)
			end if
        Next
        ShowSubFolders Subfolder
    Next
End Sub
