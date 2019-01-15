'Windowsログオン画面の景観画像をピクチャ\サンプルフォルダへコピーする

Option Explicit

dim fso
set fso = createObject("Scripting.FileSystemObject")

Dim userName
userName = CreateObject("WScript.Network").UserName

Dim assetsDir
assetsDir = "C:\Users\" & userName & "\AppData\Local\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets"

Dim tmpDir
tmpDir = "C:\Users\" & userName & "\Pictures\tmp\"
If fso.FolderExists(tmpDir) = False Then fso.CreateFolder(tmpDir)

dim outDir
outDir = "C:\Users\" & userName & "\Pictures\サンプル\"
If fso.FolderExists(outDir) = False Then fso.CreateFolder(outDir)

dim folder
set folder = fso.getFolder(assetsDir)

' ファイル一覧
dim file
dim w, h
for each file in folder.files
	dim imgName
	imgName = file.Name & ".jpg"
	fso.CopyFile file, tmpDir & imgName, True
	GetImageSize tmpDir & imgName, w, h
	
    If w > 1900 Then
        fso.CopyFile file, outDir & imgName, True
    End If
next

tmpDir = Left(tmpDir, Len(tmpDir)-1)
If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir, True

set fso = Nothing
set folder = Nothing
WScript.Echo "完了"


Sub GetImageSize(ByVal f, ByRef w, ByRef h)
    Dim strFileName, objFSO, objShellApp, objFolder, objFile
    Dim ImageSize, Width, Height
    strFileName = f

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objShellApp = CreateObject("Shell.Application")
    Set objFolder = objShellApp.Namespace(objFSO.GetParentFolderName(strFileName))
    Set objFile = objFolder.ParseName(objFSO.GetFileName(strFileName))

	'ファイルのプロパティからサイズを取得（文字列で取得されるので数値に加工）
    ImageSize = objFolder.GetDetailsOf(objFile, 31)
    Width = Split(objFolder.GetDetailsOf(objFile, 177), " ")(0)
    Width = CLng(Mid(Width, 2))
    Height = Split(objFolder.GetDetailsOf(objFile, 179), " ")(0)
    Height = CLng(Mid(Height, 2))
    If Len(ImageSize) > 0 Then
      w = Width
      h = Height
    Else
	  WScript.Echo "サイズを取得できません" & vbNewLine & f
	  Exit Sub
	End If
End Sub