// Repairer.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"

#include <iostream>
#include "atlstr.h"
#include "Shlwapi.h"

#pragma comment(lib, "Shlwapi.lib")

using namespace std;

#define FILE_7Z		_T("7z.exe")
#define FILE_DOCX	_T("std.docx")
#define FILE_XLSX	_T("std.xlsx")
#define FILE_PPTX	_T("std.pptx")
#define TEMP		_T("temp\\")

#define OPEN_TITLE	_T("请选择要修复的文件")
#if 1
#define OPEN_FILTER	TEXT("office2007 Files(*.docx;*.xlsx;*.pptx)\0*.docx;*.xlsx;*.pptx\0\0")
#else
#define OPEN_FILTER	TEXT("Word Files(*.docx)\0*.docx\0 Excel Files(*.xlsx)\0*.xlsx\0 PowerPoint Files(*.pptx)\0*.pptx\0\0")
#endif
#define MAX_BUF		40960

CString GetAppPath();
CString OpenRepairerFile();
BOOL EmptyDirectory(LPCTSTR lpszDirectory);

int _tmain(int argc, _TCHAR* argv[])
{
	CString strTemp;
	CString strRepairerFile;
	CString strParameters;

	if ( FALSE == PathFileExists(GetAppPath() + FILE_7Z) || 
		 FALSE == PathFileExists(GetAppPath() + FILE_DOCX) ||
		 FALSE == PathFileExists(GetAppPath() + FILE_XLSX) ||
		 FALSE == PathFileExists(GetAppPath() + FILE_PPTX))
	{
		// 文件不完整
		wcout<<_T("Program cann't run, please reinstall")<<endl;
		getchar();
		return -1;
	}

	strRepairerFile = OpenRepairerFile();
	if (strRepairerFile.IsEmpty())
		return -1;

	
	strTemp = GetAppPath()+TEMP;
	strParameters.Format(_T("x %s -o %s"), strRepairerFile, strTemp);
	ShellExecute(NULL, NULL, GetAppPath()+FILE_7Z, strParameters, NULL, SW_SHOWDEFAULT);

	CString strExt = strRepairerFile.Right(4);
	CString strOutFile = strRepairerFile;
	int nPos = strOutFile.ReverseFind('.');
	strOutFile.Insert(nPos, _T("_SAV"));

	if (0 == strExt.CompareNoCase(_T("xlsx")))
		CopyFile(GetAppPath() + FILE_XLSX, strOutFile, FALSE);
	else if (0 == strExt.CompareNoCase(_T("docx")))
		CopyFile(GetAppPath() + FILE_DOCX, strOutFile, FALSE);
	else if (0 == strExt.CompareNoCase(_T("pptx")))
		CopyFile(GetAppPath() + FILE_PPTX, strOutFile, FALSE);
	else
	{
		EmptyDirectory(strTemp);
		return -1;
	}

	// repaire
	strParameters.Format(_T("a %s %s*"), strOutFile, GetAppPath()+TEMP);
	ShellExecute(NULL, NULL, GetAppPath()+FILE_7Z, strParameters, NULL, SW_SHOWDEFAULT);

	wcout<<_T("Succeed!!!")<<endl;

	EmptyDirectory(strTemp);
	return 0;
}

CString GetAppPath()
{
	TCHAR szInstPath[MAX_PATH] = _T("");
	GetModuleFileName( NULL, szInstPath, MAX_PATH );
	PathRemoveFileSpec(szInstPath);
	wcscat(szInstPath, L"\\");

	return CString(szInstPath);
}

CString OpenRepairerFile()
{
	TCHAR			szBuf[MAX_BUF]; 
	OPENFILENAME	ofn;  
	
	ZeroMemory(szBuf, sizeof(szBuf));
	ZeroMemory(&ofn, sizeof(ofn));
	ofn.lStructSize = sizeof(ofn);
	ofn.lpstrFile = szBuf;
	ofn.lpstrFile[0] = _T('\0');
	ofn.nMaxFile = sizeof(szBuf);
	ofn.lpstrFilter = OPEN_FILTER;
	ofn.nFilterIndex = 1;
	ofn.lpstrFileTitle = OPEN_TITLE;
	ofn.nMaxFileTitle = 0;
	ofn.lpstrInitialDir = NULL;
	ofn.Flags = OFN_PATHMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_EXPLORER;

	if ( GetOpenFileName(&ofn) )
	{
		TCHAR szPath [MAX_PATH] = _T("");
		TCHAR szFile [MAX_PATH] = _T("");
		_tcsncpy( szPath, szBuf, ofn.nFileOffset ) ;
		szPath [ofn.nFileOffset] = _T('\0');

		int nLength = _tcslen(szPath) ;
		if (szPath[nLength - 1] != _T('\\'))
		{
			_tcscat(szPath, _T("\\"));
		}

		LPTSTR p = szBuf + ofn.nFileOffset ;
		while(*p != _T('\0'))
		{
			_tcscpy(szFile, szPath) ;
			_tcscat(szFile, p);
			p += _tcslen(p) + 1 ;
		}
		return CString(szFile) ;
	}

	return CString() ;
}

BOOL EmptyDirectory(LPCTSTR lpszDirectory)
{
	if (lstrlen(lpszDirectory)==0)
		return FALSE ;

	WIN32_FIND_DATA wfd;
	HANDLE hFind;
	TCHAR sFullPath [MAX_PATH] = _T("") ;
	TCHAR sFindFilter [MAX_PATH] = _T("") ;

	lstrcpy(sFindFilter, lpszDirectory) ;
	lstrcat(sFindFilter, _T("\\*.*")) ;

	if ((hFind = FindFirstFile(sFindFilter, &wfd)) == INVALID_HANDLE_VALUE)
		return FALSE;

	do
	{
		if (_tcscmp(wfd.cFileName, _T(".")) == 0 || _tcscmp(wfd.cFileName, _T("..")) == 0 )
			continue;

		lstrcpy(sFullPath, lpszDirectory) ;
		lstrcat(sFullPath, _T("\\")) ;
		lstrcat(sFullPath, wfd.cFileName) ;

		if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
			EmptyDirectory(sFullPath);
		else
			DeleteFile( sFullPath );

	}
	while(FindNextFile(hFind, &wfd));

	FindClose(hFind);
	//RemoveDirectory(lpszDirectory);
	return TRUE ;
}

