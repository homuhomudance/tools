'Windows���O�I����ʂ̌i�ω摜���s�N�`��\�T���v���t�H���_�փR�s�[����

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
outDir = "C:\Users\" & userName & "\Pictures\�T���v��\"
If fso.FolderExists(outDir) = False Then fso.CreateFolder(outDir)

dim folder
set folder = fso.getFolder(assetsDir)

' �t�@�C���ꗗ
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
WScript.Echo "����"


Sub GetImageSize(ByVal f, ByRef w, ByRef h)
    Dim strFileName, objFSO, objShellApp, objFolder, objFile
    Dim ImageSize, Width, Height
    strFileName = f

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objShellApp = CreateObject("Shell.Application")
    Set objFolder = objShellApp.Namespace(objFSO.GetParentFolderName(strFileName))
    Set objFile = objFolder.ParseName(objFSO.GetFileName(strFileName))

	'�t�@�C���̃v���p�e�B����T�C�Y���擾�i������Ŏ擾�����̂Ő��l�ɉ��H�j
    ImageSize = objFolder.GetDetailsOf(objFile, 31)
    Width = Split(objFolder.GetDetailsOf(objFile, 177), " ")(0)
    Width = CLng(Mid(Width, 2))
    Height = Split(objFolder.GetDetailsOf(objFile, 179), " ")(0)
    Height = CLng(Mid(Height, 2))
    If Len(ImageSize) > 0 Then
      w = Width
      h = Height
    Else
	  WScript.Echo "�T�C�Y���擾�ł��܂���" & vbNewLine & f
	  Exit Sub
	End If
End Sub