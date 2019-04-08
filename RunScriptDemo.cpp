// RunScriptDemo.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"
#include "SimpleScriptSite.h"
#include <iostream>
#include <stdio.h>

void testExpression(const wchar_t* prefix, IActiveScriptParse* pScriptParse, LPCOLESTR script)
{
	wprintf(L"%s: %s: ", prefix, script);
	HRESULT hr = S_OK;
	CComVariant result;
	EXCEPINFO ei = { };
	hr = pScriptParse->ParseScriptText(script, NULL, NULL, NULL, 0, 0, SCRIPTTEXT_ISEXPRESSION, &result, &ei);
	CComVariant resultBSTR;
	VariantChangeType(&resultBSTR, &result, 0, VT_BSTR);
	wprintf(L"%s\n", V_BSTR(&resultBSTR));
}

void testScript(const wchar_t* prefix, IActiveScriptParse* pScriptParse, LPCOLESTR script)
{
	wprintf(L"%s: %s\n", prefix, script);
	HRESULT hr = S_OK;
	CComVariant result;
	EXCEPINFO ei = { };
	hr = pScriptParse->ParseScriptText(script, NULL, NULL, NULL, 0, 0, 0, &result, &ei);
}

int _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hr = S_OK;
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	// Initialize
	CSimpleScriptSite* pScriptSite = new CSimpleScriptSite();
	CComPtr<IActiveScript> spJScript;
	CComPtr<IActiveScriptParse> spJScriptParse;
	hr = spJScript.CoCreateInstance(OLESTR("JScript"));
	hr = spJScript->SetScriptSite(pScriptSite);
	hr = spJScript->QueryInterface(&spJScriptParse);
	hr = spJScriptParse->InitNew();
	CComPtr<IActiveScript> spVBScript;
	CComPtr<IActiveScriptParse> spVBScriptParse;
	hr = spVBScript.CoCreateInstance(OLESTR("VBScript"));
	hr = spVBScript->SetScriptSite(pScriptSite);
	hr = spVBScript->QueryInterface(&spVBScriptParse);
	hr = spVBScriptParse->InitNew();

	// Run some scripts
	testExpression(L"JScript", spJScriptParse, OLESTR("(new Date).getTime()"));
	testExpression(L"VBScript", spVBScriptParse, OLESTR("Now"));
	testScript(L"VBScript", spVBScriptParse, OLESTR("MsgBox \"Hello World! The current time is: \" & Now"));
    // Outputs:
    // JScript: (new Date).getTime() : 1554685325378
	// VScript: Now : 8 / 04 / 2019 11 : 02 : 05 AM
	// VBScript: MsgBox "Hello World! The current time is: " & Now

	// Cleanup
	spVBScriptParse = NULL;
	spVBScript = NULL;
	spJScriptParse = NULL;
	spJScript = NULL;
	pScriptSite->Release();
	pScriptSite = NULL;

	::CoUninitialize();
	return 0;
}