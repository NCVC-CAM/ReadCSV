// ReadCSV.cpp : DLL 用の初期化処理の定義を行います。
//

#include "stdafx.h"
#include <afxdllx.h>
#include "resource.h"
#include "NCVCaddin.h"

#include <math.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static	const	char*	g_szFileExt[] = {"csv"};
static	const	char*	g_szSerialFunc = "ReadCSV";

const double PI = 3.1415926535897932384626433832795;
const double RAD = PI / 180.0;
const double DEG = 180.0 / PI;

static AFX_EXTENSION_MODULE ReadCSVDLL = { NULL, NULL };

extern "C" int APIENTRY
DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	// lpReserved を使う場合はここを削除してください
	UNREFERENCED_PARAMETER(lpReserved);

	if (dwReason == DLL_PROCESS_ATTACH)
	{
		TRACE0("READCSV.DLL Initializing!\n");
		
		// 拡張 DLL を１回だけ初期化します。
		if (!AfxInitExtensionModule(ReadCSVDLL, hInstance))
			return 0;

		// この DLL をリソース チェインへ挿入します。
		// メモ: 拡張 DLL が MFC アプリケーションではなく
		//   MFC 標準 DLL (ActiveX コントロールのような)
		//   に暗黙的にリンクされる場合、この行を DllMain
		//   から削除して、この拡張 DLL からエクスポート
		//   された別の関数内へ追加してください。  
		//   この拡張 DLL を使用する標準 DLL はこの拡張 DLL
		//   を初期化するために明示的にその関数を呼び出します。 
		//   それ以外の場合は、CDynLinkLibrary オブジェクトは
		//   標準 DLL のリソース チェインへアタッチされず、
		//   その結果重大な問題となります。

		new CDynLinkLibrary(ReadCSVDLL);
	}
	else if (dwReason == DLL_PROCESS_DETACH)
	{
		TRACE0("READCSV.DLL Terminating!\n");
		// デストラクタが呼び出される前にライブラリを終了します
		AfxTermExtensionModule(ReadCSVDLL);
	}
	return 1;   // ok
}

/////////////////////////////////////////////////////////////////////////////
// NCVC ｱﾄﾞｲﾝ関数

NCADDIN BOOL NCVC_Initialize(NCVCINITIALIZE* nci)
{
	static	const	char	szReadMenuName[] = "CSVﾃﾞｰﾀを開く...";
	static	const	char	szReadFuncName[] = "ReadCSV_Menu";
	static	const	int		nReadMenu[] = {
		NCVCADIN_ARY_APPFILE, NCVCADIN_ARY_NCDFILE, NCVCADIN_ARY_DXFFILE
	};
/*
	static	const	char	szOptMenuName[] = "CSVﾃﾞｰﾀのｵﾌﾟｼｮﾝ設定...";
	static	const	char	szOptFuncName[] = "ReadCSV_Opt";
	static	const	int		nOptMenu[] = {
		NCVCADIN_ARY_APPOPTION, NCVCADIN_ARY_NCDOPTION, NCVCADIN_ARY_DXFOPTION
	};
*/
	int		i;

	// ｱﾄﾞｲﾝの必要情報
	nci->dwSize = sizeof(NCVCINITIALIZE);
	nci->dwType = NCVCADIN_FLG_APPFILE|NCVCADIN_FLG_NCDFILE|NCVCADIN_FLG_DXFFILE;
//		NCVCADIN_FLG_APPOPTION|NCVCADIN_FLG_NCDOPTION|NCVCADIN_FLG_DXFOPTION;
//	nci->nToolBar = 0;
	for ( i=0; i<SIZEOF(nReadMenu); i++ ) {
		nci->lpszMenuName[nReadMenu[i]] = szReadMenuName;
		nci->lpszFuncName[nReadMenu[i]] = szReadFuncName;
	}
/*
	for ( i=0; i<SIZEOF(nOptMenu); i++ ) {
		nci->lpszMenuName[nOptMenu[i]] = szOptMenuName;
		nci->lpszFuncName[nOptMenu[i]] = szOptFuncName;
	}
*/
	nci->lpszAddinName	= g_szSerialFunc;
	nci->lpszCopyright	= "MNCT Yoshiro Nose";
	nci->lpszSupport	= "http://s-gikan2.maizuru-ct.ac.jp/";

	// ｶｽﾀﾑ拡張子の登録．これをしておくとD&Dでもﾌｧｲﾙが開ける
	for ( i=0; i<SIZEOF(g_szFileExt); i++ )
		NCVC_AddDXFExtensionFunc(g_szFileExt[i], g_szSerialFunc, g_szSerialFunc);

	return TRUE;
}


/////////////////////////////////////////////////////////////////////////////
// NCVC ｲﾍﾞﾝﾄ関数

NCADDIN void ReadCSV_Menu(void)
{
	CFileDialog	dlg(TRUE, g_szFileExt[0], NULL, OFN_FILEMUSTEXIST|OFN_HIDEREADONLY,
		"CSV files (*.csv)|*.csv|Text file (*.txt)|*.txt|All Files (*.*)|*.*||");
	if ( dlg.DoModal() != IDOK )
		return;
	// 新規ﾄﾞｷｭﾒﾝﾄの作成
	NCVC_CreateDXFDocument(dlg.GetPathName(), g_szSerialFunc);
/*
	// ﾌﾟﾚｰｽﾊﾞｰ表示のためSDK流に書き直し
	TCHAR			szFileName[_MAX_PATH];
	::ZeroMemory(szFileName, sizeof(_MAX_PATH));
	OPENFILENAME	ofn;
	::ZeroMemory(&ofn, sizeof(OPENFILENAME));
	ofn.lStructSize	= sizeof(OPENFILENAME);
	ofn.hwndOwner	= NCVC_GetMainWnd();
	ofn.lpstrFilter	= "CSV files (*.csv)\0*.csv\0TXT file (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0";
	ofn.lpstrFile	= szFileName;
	ofn.nMaxFile	= _MAX_PATH;
	ofn.Flags		= OFN_FILEMUSTEXIST|OFN_HIDEREADONLY;
	ofn.lpstrDefExt	= g_szFileExt[0];
	if ( !::GetOpenFileName(&ofn) )
		return;

	CString	strFile(ofn.lpstrFile);
	// 新規ﾄﾞｷｭﾒﾝﾄの作成
	NCVC_CreateDXFDocument(strFile, g_szSerialFunc);
*/
}
/*
NCADDIN void ReadCSV_Opt(void)
{
	AfxMessageBox("Option");
}
*/
/////////////////////////////////////////////////////////////////////////////
//	関数プロトタイプ
static BOOL GetCsvPoint(NCVCHANDLE, LPCTSTR, double[]);
static BOOL GetCsvLine(NCVCHANDLE, LPCTSTR, double[]);
static BOOL GetCsvCircle(NCVCHANDLE, LPCTSTR, int, double[]);

/////////////////////////////////////////////////////////////////////////////
//	JWC 読み込み
//	NCVC_CreateDXFDocument() の第２引数で指定．
//	NCVC本体から自動的に呼ばれる

NCADDIN BOOL ReadCSV(NCVCHANDLE hDoc, LPCTSTR pszFile)
{
//	AfxMessageBox("ReadCSV Read Start!!");

	static	LPCTSTR		szDelimiter = ", \t";

	CString	strBuf, strTmp, strLayer;
	double	nData[8];
	int		nNumOfData, nNumOfErr = 0, nAlloc = 0;
	char*	szTok;
	char*	szContext;
	char*	szBuf = NULL;
	BOOL	bResult = TRUE, bOrigin = FALSE;	// 原点ﾚｲﾔ処理中か

	try {
		CStdioFile	fp(pszFile, CFile::modeRead | CFile::shareDenyWrite);

		while ( fp.ReadString(strBuf) ) {				// １行ずつ読み込み
			// char* へｺﾋﾟｰ
			if ( strBuf.GetLength() >= nAlloc ) {
				if ( szBuf )
					delete	szBuf;
				nAlloc = strBuf.GetLength();
				szBuf = new char[nAlloc+1];
			}
			lstrcpy(szBuf, strBuf);
			nNumOfData = 0;	// データ数初期化
			if ( strBuf.FindOneOf(szDelimiter) >= 0 ) {				// szDelimeter が見つかれば
				szTok = strtok_s(szBuf, szDelimiter, &szContext);		// szDelimeter を区切りとしてﾄｰｸﾝ取りだし
				while ( szTok != NULL && nNumOfData < SIZEOF(nData) ){
					nData[nNumOfData] = atof(szTok);
					szTok = strtok_s(NULL, szDelimiter, &szContext);
					nNumOfData++;
				}
			}
			else
				szTok = szBuf;							// そうでなければ１行丸ごと szTok へ

			if ( nNumOfData!=0 && strLayer.IsEmpty() )	// ｶｳﾝﾄ!=0 かつ ﾚｲﾔ名が空なら while ﾙｰﾌﾟの先頭へ
				continue;

			switch(nNumOfData){
			case 0:		// ﾚｲﾔデータ
				bOrigin = NCVC_IsOriginLayer(szTok) ? TRUE : FALSE;
				if ( bOrigin || NCVC_IsCutterLayer(szTok) )
					strLayer = szTok;	// szTok が NCVC で設定した切削ﾚｲﾔか原点ﾚｲﾔならﾚｲﾔ名に代入
				else
					strLayer.Empty();
				break;
			case 2:		// 点データ
				if (!GetCsvPoint(hDoc, strLayer, nData)) return FALSE;
				break;
			case 3:		// 円データ
				if ( bOrigin ) {
					// 原点ﾃﾞｰﾀ登録
					DPOINT	pt;
					pt.x = nData[0];
					pt.y = nData[1];
					NCVC_SetDXFCutterOrigin(hDoc, &pt, nData[2], FALSE);
				}
				else if (!GetCsvCircle(hDoc, strLayer, 0, nData)) return FALSE;
				break;
			case 4:		// 線データ
				if (!GetCsvLine(hDoc, strLayer, nData)) return FALSE;
				break;
			case 5:		// 円弧データ
				if (!GetCsvCircle(hDoc, strLayer, 1, nData)) return FALSE;
				break;
			case 7:		// 楕円データ
				if (!GetCsvCircle(hDoc, strLayer, 2, nData)) return FALSE;
				break;
			default:	// 無効なデータ
				nNumOfErr++;
				break;
			}	// switch
		}	// while(main)
	}	// try
	catch ( CMemoryException* e ) {
		AfxMessageBox("Out of memory.", MB_OK|MB_ICONSTOP);
		e->Delete();
		bResult = FALSE;
	}
	catch (	CFileException* e ) {
		AfxMessageBox("ﾌｧｲﾙｱｸｾｽで例外が発生しました", MB_OK|MB_ICONSTOP);
		e->Delete();
		bResult = FALSE;
	}
	
	if ( szBuf )
		delete	szBuf;
	
	if ( nNumOfErr > 0 ){
		strTmp.Format("%d 個の無効なデータがありました。", nNumOfErr);
		AfxMessageBox(strTmp);
	}

	return bResult;
}


/////////////////////////////////////////////////////////////////////////////
// 各種データ解析関数

//// 点データ処理
BOOL GetCsvPoint(NCVCHANDLE hDoc, LPCTSTR lpszLayer, double nData[]){
	DXFDATA		dxf;

	dxf.dwSize = sizeof(DXFDATA);
	dxf.enType = DXFPOINTDATA;

	lstrcpy(dxf.szLayer, lpszLayer);
	dxf.nLayer = DXFCAMLAYER;
	dxf.ptS.x = nData[0];
	dxf.ptS.y = nData[1];

	if ( !NCVC_AddDXFData(hDoc, &dxf) ){
		AfxMessageBox("点データの読み込みに失敗しました。");
		return FALSE;
	}

	return TRUE;
}

//// 線データ処理
BOOL GetCsvLine(NCVCHANDLE hDoc, LPCTSTR lpszLayer, double nData[]){
	DXFDATA		dxf;

	dxf.dwSize = sizeof(DXFDATA);
	dxf.enType = DXFLINEDATA;

	lstrcpy(dxf.szLayer, lpszLayer);
	dxf.nLayer = DXFCAMLAYER;
	dxf.ptS.x = nData[0];
	dxf.ptS.y = nData[1];
	dxf.de.ptE.x = nData[2];
	dxf.de.ptE.y = nData[3];

	if ( !NCVC_AddDXFData(hDoc, &dxf) ){
		AfxMessageBox("線データの読み込みに失敗しました。");
		return FALSE;
	}

	return TRUE;
}

//// 円データ処理
BOOL GetCsvCircle(NCVCHANDLE hDoc, LPCTSTR lpszLayer, int CircleType, double nData[]){
	DXFDATA		dxf;
	double		lq;

	// 円ﾃﾞｰﾀ代入(共通)
	dxf.dwSize = sizeof(DXFDATA);
	dxf.ptS.x = nData[0];
	dxf.ptS.y = nData[1];
	
	lstrcpy(dxf.szLayer, lpszLayer);
	dxf.nLayer = DXFCAMLAYER;
	switch ( CircleType ){
	case 0:		// 円
		dxf.enType = DXFCIRCLEDATA;
		dxf.de.dR  = nData[2];
		break;
	case 1:		// 円弧（角度単位は度）
		dxf.enType = DXFARCDATA;
		dxf.de.arc.r  = nData[2];
		dxf.de.arc.sq = nData[3];
		dxf.de.arc.eq = nData[4];
		break;
	case 2:		// 楕円（弧）（角度単位は度）
		dxf.enType = DXFELLIPSEDATA;
		dxf.de.elli.sq = nData[3] * RAD;
		dxf.de.elli.eq = nData[4] * RAD;
		lq = nData[6] * RAD;
		dxf.de.elli.ptL.x = nData[2] * cos(lq);
		dxf.de.elli.ptL.y = nData[2] * sin(lq);
		dxf.de.elli.s = nData[5];
		break;
	default:
		break;
	}

	if ( !NCVC_AddDXFData(hDoc, &dxf) ){
		AfxMessageBox("円(弧)または楕円データの読み込みに失敗しました。");
		return FALSE;
	}

	return TRUE;
}
