/*+===================================================================
  File:      main.cpp

  Summary:   Data import GUI for the solarTemperatureLogger project (Developed with Ebmarcadero DEV-C++ 6.3)

  License: 2-Clause BSD License
  Copyright (c) 2024 codingABI

  Icons for the toolbar:
  	Icons for Import, Save, Information and Exit: Game icon pack by Kenney Vleugels (www.kenney.nl), https://kenney.nl/assets/game-icons, CC0
 	Icon for Refresh: https://commons.wikimedia.org/wiki/File:Refresh_icon.svg, CC0

  History:
  20240717, Initial version
  20240718, Add simple graph for data values
  20240805, Resize window when dpi have changed
  20240913, Fix spelling for "Mouseubttons"
  20241113, Restore SelectObject in paintGraph
  20251002, Add WINAPI for MyGet* functions to support x86 binaries

===================================================================+*/

#define UNICODE

// Set _WIN32_WINNT to minimum level Windows 8
#define _WIN32_WINNT 0x0602

#define _WIN32_DCOM

#include <windows.h>
#include <regex>
#include <commctrl.h>
#include <cmath>
#include <algorithm>
#include <VersionHelpers.h>
#include <shlwapi.h>
#include <comdef.h>
#include <Wbemidl.h>
#include <Dbt.h>

#include "resource.h"
#include "resourceMui.h"

#define WORKINGCSVFILE L"solarTemperatureLoggerTool.csv" // Internal CSV data file
#define BAUDRATE 9600 // Baud rate for the serial adapter
#define TOOLBARMAXITEMS 5 // Max items in toolbar
#define _MAX_ITOSTR_BASE16_COUNT (8 + 1) // Char length for DWORD to hex conversion 

// Return codes for the serial adapter rescan thread
#define SA_OK 0
#define SA_ERROR 1

// Default sizes for some controls in 96 dpi
#define STATUSBARLEFTWIDTH_96DPI 150
#define TOOLBARSEPWIDTH_96DPI 6
#define INITIALTABLELEFTWIDTH_96DPI 150
#define WINDOWMINWIDTH_96DPI 400
#define WINDOWMINHEIGHT_96DPI 200
#define GRAPHHEIGHT_96DPI 50

// Global variables
int g_iFontHeight_96DPI = -12;
HINSTANCE g_hInst = NULL;
HFONT g_hFont = NULL;
HIMAGELIST g_hToolbarImageList16 = NULL;
HIMAGELIST g_hToolbarImageList32 = NULL;
HANDLE g_hThreadRescanSerialAdapter = NULL;
HANDLE g_hThreadImportSerialData = NULL;
HANDLE g_hSerialHandle = NULL;
HANDLE g_hImportAbortEvent = NULL;
HMENU g_hMenu = NULL;
HMENU g_hContextMenu = NULL;
HWND g_hTable = NULL;
HWND g_hGraph = NULL;
HWND g_hEventList = NULL;
HWND g_hComboBox = NULL;
HANDLE g_hSemaphoreDatasets = NULL;
WNDPROC g_WindowProcTableOrig = NULL;
WNDPROC g_WindowProcGraphOrig = NULL;
int g_iLastSelectedCOM = -1;

// Dataset vector (used for graph)
typedef struct _DATASET {
	SYSTEMTIME stUTC;
	double temperature;
} DATASET;
std::vector<DATASET> g_vDatasets;

/*
 * I had a problem with the built in GetDpiForWindow function.
 * When I use this function and run the program on Windows 2012R2 the program does not start and I will get the message
 * 'Der Prozedureinsprungpunkt "GetDpiForWindow" wurde in der DLL "...exe" nicht gefunden. '
 * a.k.a. 'The procedure entry point "GetDpiForWindow" could not be located in the dynamic link library "...exe"'
 * My workaround is MyGetDpiForWindow
 */
UINT MyGetDpiForWindow(HWND hWindow) {
	UINT uDpi = 96; // Failback value, if GetDpiForWindow could not be used

	UINT (WINAPI* GetDpiForWindow)(HWND hwnd) = nullptr;
    HMODULE hDLL = GetModuleHandle(L"user32.dll");
    if (hDLL == NULL)
        return uDpi;
    *reinterpret_cast<FARPROC*>(&GetDpiForWindow) = GetProcAddress(hDLL, "GetDpiForWindow");
    if (GetDpiForWindow == nullptr)
        return uDpi;
 	return(GetDpiForWindow(hWindow));
}

/*
 * I had a problem with the built in GetSystemMetricsForDpi function.
 * When I use this function and run the program on Windows 2012R2 the program does not start and I will get the message
 * 'Der Prozedureinsprungpunkt "GetSystemMetricsForDpi" wurde in der DLL "...exe" nicht gefunden. '
 * a.k.a. 'The procedure entry point "GetSystemMetricsForDpi" could not be located in the dynamic link library "...exe"'
 * My workaround is MyGetSystemMetricsForDpi
 */
int MyGetSystemMetricsForDpi(int nIndex, UINT dpi) {
	int iMetrics = GetSystemMetrics(nIndex); // Failback value, if GetSystemMetricsForDpi could not be used

	int (WINAPI* GetSystemMetricsForDpi)(int nIndex, UINT dpi) = nullptr;
    HMODULE hDLL = GetModuleHandle(L"user32.dll");
    if (hDLL == NULL)
        return iMetrics;
    *reinterpret_cast<FARPROC*>(&GetSystemMetricsForDpi) = GetProcAddress(hDLL, "GetSystemMetricsForDpi");
    if (GetSystemMetricsForDpi == nullptr)
        return iMetrics;
 	return(GetSystemMetricsForDpi(nIndex, dpi));
}

// From Reactos comsupp.cpp to get https://learn.microsoft.com/en-us/windows/win32/wmisdk/example--getting-wmi-data-from-the-local-computer running on Dev-C++
namespace _com_util
{

BSTR WINAPI ConvertStringToBSTR(const char *pSrc)
{
    DWORD cwch;
    BSTR wsOut(NULL);

    if (!pSrc) return NULL;

    /* Compute the needed size with the NULL terminator */
    cwch = ::MultiByteToWideChar(CP_ACP, 0, pSrc, -1, NULL, 0);
    if (cwch == 0) return NULL;

    /* Allocate the BSTR (without the NULL terminator) */
    wsOut = ::SysAllocStringLen(NULL, cwch - 1);
    if (!wsOut)
    {
        ::_com_issue_error(HRESULT_FROM_WIN32(ERROR_OUTOFMEMORY));
        return NULL;
    }

    /* Convert the string */
    if (::MultiByteToWideChar(CP_ACP, 0, pSrc, -1, wsOut, cwch) == 0)
    {
        /* We failed, clean everything up */
        cwch = ::GetLastError();

        ::SysFreeString(wsOut);
        wsOut = NULL;

        ::_com_issue_error(!IS_ERROR(cwch) ? HRESULT_FROM_WIN32(cwch) : cwch);
    }

    return wsOut;
}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: LoadStringAsWstr

  Summary:  Get resource string as wstring

  Args:     HINSTANCE hInstance
  			  Handle to instance module
  			UINT uID
  			  Resource ID

  Returns:  std::wstring

-----------------------------------------------------------------F-F*/
std::wstring LoadStringAsWstr(HINSTANCE hInstance, UINT uID) { 
	PCWSTR pws; 
	int cchStringLength = LoadStringW(hInstance, uID, reinterpret_cast<LPWSTR>(&pws), 0);
	if (cchStringLength > 0) return std::wstring(pws, cchStringLength); else return std::wstring();
}


/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: ErrorMessage

  Summary:  Show message box with an error icon, a message, error code and error message

  Args:     HWND hWindow
			  Window handle
  			std::wstring sMessage
  			  Message
			DWORD dwError
              Error code from GetLastError

  Returns:  BOOL
              TRUE = Success
			  FALSE = Error
-----------------------------------------------------------------F-F*/
BOOL ErrorMessage(HWND hWindow,std::wstring sMessage, DWORD dwError) {
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT+2];
	std::wstring sErrorMessage = L"";
    LPWSTR messageBuffer = nullptr;

	if (hWindow == NULL) {
		MessageBox(NULL, L"Window handle is NULL",L"Error",MB_OK|MB_ICONERROR);
		return FALSE;
	}

	if (dwError > 0) {
		// Convert error to hex
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", dwError);
		// Get message for error
	    size_t size = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
	        NULL, dwError, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&messageBuffer, 0, NULL);
		sErrorMessage.assign(L"\nError code ").append(szHex).append(L"\n").append(messageBuffer, size);
	}

	MessageBox(hWindow, sMessage.append(sErrorMessage).c_str(), LoadStringAsWstr(g_hInst,IDS_MSGBOXERRORTITLE).c_str(),MB_OK|MB_ICONERROR);
	// Free buffer
    LocalFree(messageBuffer);
	return TRUE;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: openWithFileExplorer

  Summary:  Open the dataset folder in the file explorer

  Args:     HWND hWindow
			  Window handle

  Returns:
-----------------------------------------------------------------F-F*/
void openWithFileExplorer(HWND hWindow,const wchar_t *pszPath ) {
	if (hWindow == NULL) return;
	if (pszPath == NULL) return;
	ShellExecute(hWindow, L"open", pszPath, NULL, NULL, SW_SHOWNORMAL);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: setDisabledState

  Summary:  Append message to the event list

  Args:     const wchar_t *pszMessage
              Pointer to message string

  Returns:
-----------------------------------------------------------------F-F*/
void addMessageToEventList(const wchar_t *pszMessage) {
	if (g_hEventList == NULL) return;
	if (pszMessage == NULL) return;
	int iCount = SendMessage(g_hEventList, LB_GETCOUNT, 0, (LPARAM) 0);
   	SendMessage(g_hEventList, LB_ADDSTRING, 0, (LPARAM) pszMessage);
   	SendMessage(g_hEventList, LB_SETTOPINDEX, (WPARAM) iCount,0); // "Scroll down"
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: isRunningUnderWine

  Summary:  Check if program is running under wine

  Args:

  Returns:  BOOL
              TRUE = Yes, running on wine
			  FALSE = No, not running on wine or not running under window nt
-----------------------------------------------------------------F-F*/
BOOL isRunningUnderWine() {
    HMODULE hDLL = GetModuleHandle(L"ntdll.dll");
    if(hDLL == NULL) return FALSE;
    if (GetProcAddress(hDLL, "wine_get_version") == NULL) return FALSE; else return TRUE;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: setDisabledState

  Summary:  Set disabled state for toolbar and menu item

  Args:     HWND hWindow
              Handle to main window
            int nIDDlgItem
              ID of GUI control
            bool bState
              bState = false => Enable item, bState = true => Disable item

  Returns:  BOOL
              TRUE = Success
			  FALSE = Error
-----------------------------------------------------------------F-F*/
BOOL setDisabledState(HWND hWindow, int nIDDlgItem, bool bState) {
	if (hWindow == NULL) {
		MessageBox(NULL, L"Window handle is NULL",L"Error",MB_OK|MB_ICONERROR);
		return FALSE;
	}
	HWND hToolBar = GetDlgItem(hWindow, IDC_TOOLBAR);
	if (hToolBar ==  NULL) {
		addMessageToEventList(L"Error finding tool bar");
		return FALSE;
	}
	SendMessage(hToolBar, TB_ENABLEBUTTON , nIDDlgItem, !bState);
	EnableMenuItem(g_hMenu, nIDDlgItem,bState);
	return TRUE;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: SerialInit

  Summary:  Setup serial connection

  Args:     HWND hWindow
              Handle to main window
            const wchar_t *ComPortName
              Serial device name
            int BaudRate
              Baud rate

  Returns:  HANDLE
              Handle to serial adapter.
			  NULL, if function fails
-----------------------------------------------------------------F-F*/
HANDLE SerialInit(HWND hWindow, const wchar_t *ComPortName, int BaudRate)
{
	DCB dcb;
	COMMTIMEOUTS CommTimeouts;
	
	#define FC_DTRDSR 0x01
	#define FC_RTSCTS 0x02
	#define FC_XONXOFF 0x04
	
	#define ASCII_BEL 0x07
	#define ASCII_BS 0x08
	#define ASCII_LF 0x0A
	#define ASCII_CR 0x0D
	#define ASCII_XON 0x11
	#define ASCII_XOFF 0x13
	
	// Flow control flags
	
	#define FC_DTRDSR 0x01
	#define FC_RTSCTS 0x02
	#define FC_XONXOFF 0x04
	
	// ascii definitions
	
	#define ASCII_BEL 0x07
	#define ASCII_BS 0x08
	#define ASCII_LF 0x0A
	#define ASCII_CR 0x0D
	#define ASCII_XON 0x11
	#define ASCII_XOFF 0x13
	
	HANDLE hCom = INVALID_HANDLE_VALUE;
	bool bPortReady;
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT+2];
	std::wstring sMessage;

	hCom = CreateFile(ComPortName,
		GENERIC_READ | GENERIC_WRITE,
		0, // No shared access
		NULL, // No security
		OPEN_EXISTING,
		0, // No overlapped I/O
		NULL); // No template
	if (hCom == INVALID_HANDLE_VALUE)  {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"Error opening COM. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
		return(NULL);
	}

	bPortReady = SetupComm(hCom, 2, 128); // Set buffer sizes
	if (!bPortReady)  {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"Error setup comm. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
		CloseHandle(hCom);
		return(NULL);
	}

	bPortReady = GetCommState(hCom, &dcb);
	if (!bPortReady)  {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"Error getting comm state. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
		CloseHandle(hCom);
		return(NULL);
	}

	dcb.BaudRate = BaudRate;
	dcb.ByteSize = 8;
	dcb.Parity = NOPARITY;
	dcb.StopBits = ONESTOPBIT;
	dcb.fAbortOnError = TRUE;

	// Set XON/XOFF
	dcb.fOutX = FALSE; // XON/XOFF off for transmit
	dcb.fInX = FALSE; // XON/XOFF off for receive
	// Set RTSCTS
	dcb.fOutxCtsFlow = FALSE; // TRUE; // turn on CTS flow control
	dcb.fRtsControl = RTS_CONTROL_HANDSHAKE; //
	// Set DSRDTR
	dcb.fOutxDsrFlow = FALSE; // turn on DSR flow control
	dcb.fDtrControl = DTR_CONTROL_ENABLE; //

	bPortReady = SetCommState(hCom, &dcb);
	if (!bPortReady)  {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"Error setting comm state. Fake CH340?. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
		CloseHandle(hCom);
		return(NULL);
	}

	// Communication timeouts are optional
	bPortReady = GetCommTimeouts (hCom, &CommTimeouts);
	if (!bPortReady)  {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"Error getting comm timeouts. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
		CloseHandle(hCom);
		return(NULL);
	}

	CommTimeouts.ReadIntervalTimeout = 5000;
	CommTimeouts.ReadTotalTimeoutConstant = 5000;
	CommTimeouts.ReadTotalTimeoutMultiplier = 1000;
	CommTimeouts.WriteTotalTimeoutConstant = 5000;
	CommTimeouts.WriteTotalTimeoutMultiplier = 1000;

	bPortReady = SetCommTimeouts (hCom, &CommTimeouts);
	if (!bPortReady)  {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"Error setting comm timeouts. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
		CloseHandle(hCom);
		return(NULL);
	}

	return hCom;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: threadImportSerialData

  Summary:  Thread for waiting for serial data

  Args:     void *data
              Pointer to HWND of main window

  Returns:  unsigned int
              SF_COMERROR = Serial error
			  SI_FILEERROR = File error
			  SF_OK = Success
			  SF_TIMEOUT = Timeout expired
			  SI_ABORT = Aborted by user request

-----------------------------------------------------------------F-F*/
unsigned int __stdcall threadImportSerialData(void *data) {
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT+2];
	std::wstring sMessage;
	#define MAXLINELENGHT 254
	#define MAXWAITTIME 30 // Timeout in seconds
	std::wstring sCOMport = L"";
	HANDLE serialHandle;
	std::wstring sData = L"\ufeff"; // Needed unicode BOM to get a UTF 16 big endian file 
	std::wstring sLine = L"";
	bool exitLoop = false;
	bool beginDetected = false;
	clock_t startTime = 0;
	char rxchar;
	BOOL bReadRC;
	DWORD iBytesRead;
	static int iLastProgessBar = -100;

	HWND *phWindow = (HWND*) data;
	HWND hWindow = *phWindow;
	free(phWindow);

    if (hWindow == NULL) { // Should not happen
		CloseHandle(g_hSerialHandle);
		g_hSerialHandle = NULL;
		MessageBox(NULL, L"Window handle is NULL",L"Error",MB_OK|MB_ICONERROR);
    	return SI_COMERROR;
	}

	sMessage.assign(L"Wait for serial data. Thread ID ").append(std::to_wstring(GetCurrentThreadId()));
	addMessageToEventList(sMessage.c_str());

	ResetEvent(g_hImportAbortEvent); // Reset pending events
	HWND hStatusBar = GetDlgItem(hWindow, IDC_STATUSBAR);
	HWND hProgessBar = GetDlgItem(hWindow, IDC_PROGRESSBAR);
	HWND hProgressButton = GetDlgItem(hWindow, IDC_PROGRESSBUTTON);
	HWND hProgressText = GetDlgItem(hWindow, IDC_PROGRESSTEXT);

	if ((hStatusBar == NULL) || (hProgessBar == NULL) || (hProgressButton == NULL) || (hProgressText == NULL)) {
		CloseHandle(g_hSerialHandle);
		g_hSerialHandle = NULL;
		addMessageToEventList(L"Error finding controls");
    	return SI_COMERROR;
	}

    SendMessage(hProgessBar, PBM_SETPOS, (WPARAM) 100, 0); // Set progress bar to 100%
	ShowWindow(hStatusBar, SW_HIDE);
	ShowWindow(hProgessBar, SW_SHOW);
	EnableWindow(hProgressButton,true);
	ShowWindow(hProgressButton, SW_SHOW);
	SendMessage(hProgressText, WM_SETTEXT, (WPARAM)0, (LPARAM) LoadStringAsWstr(g_hInst,IDS_WAITINGFORDATA).c_str());
	ShowWindow(hProgressText, SW_SHOW);

	startTime = clock();
	int iRC = SI_OK;
	do {
		bReadRC = ReadFile(g_hSerialHandle, &rxchar, 1, &iBytesRead, NULL);

		if (iBytesRead == 0) {
			DWORD lastError = GetLastError();
			if (GetLastError() == 995) { // Sometimes first connect fail => ClearCommError and retry
				DWORD dwErrorFlags;
				COMSTAT ComStat;
				ClearCommError(g_hSerialHandle,&dwErrorFlags,&ComStat);
				rxchar = '\0';
				sLine=L"";
 			} else if (GetLastError() != 0) {
			 	_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
				sMessage.assign(L"Error reading port. Error code ").append(szHex);
				addMessageToEventList(sMessage.c_str());
				iRC = SI_COMERROR;
			}
		} else {

			if ((rxchar == '\r') || (rxchar == '\n'))  {
				if (sLine == L"END") {
					exitLoop = true;
				} else {
					if (sLine == L"BEGIN") {
						beginDetected = true;
						sLine = L"";
					} else {
						if (beginDetected && (sLine.length()>0)) {
							sData.append(sLine).append(L"\n");
						}
					}
				}
				sLine = L"";
				rxchar = '\0';
			} else {
				if (sLine.length() < MAXLINELENGHT) {
					if ((rxchar != '\r') && (rxchar != '\n')) sLine+=rxchar;
				}
			}
		}

		int iProgressBar =  100-(100.0f*((float)((clock()-startTime)/CLOCKS_PER_SEC)/MAXWAITTIME));
		if (iProgressBar < 0) iProgressBar = 0;
		if (iProgressBar > 100) iProgressBar = 100;
		if (iProgressBar != iLastProgessBar) {
	    	SendMessage(hProgessBar, PBM_SETPOS, (WPARAM) iProgressBar, 0);
			iLastProgessBar = iProgressBar;
		}

		if ((iRC == SI_OK) && (!exitLoop) && (iProgressBar == 0)) iRC = SI_TIMEOUT;
		if (WaitForSingleObject(g_hImportAbortEvent, 0) ==  WAIT_OBJECT_0) iRC = SI_ABORT;
		if (iRC != SI_OK) exitLoop = true;
	} while (!exitLoop);


	ShowWindow(hProgessBar, SW_HIDE);
	ShowWindow(hProgressButton, SW_HIDE);
	ShowWindow(hProgressText, SW_HIDE);
	ShowWindow(hStatusBar, SW_SHOW);

	CloseHandle(g_hSerialHandle);
	g_hSerialHandle = NULL;

	wchar_t szWorkingFolder[MAX_PATH];
	GetModuleFileName(NULL,szWorkingFolder,MAX_PATH);
	PathRemoveFileSpec(szWorkingFolder);
	std::wstring sFullPathWorkingFile;
	sFullPathWorkingFile.assign(szWorkingFolder).append(L"\\").append(WORKINGCSVFILE);

	if (iRC == SI_OK) {
		// Open CSV file
		HANDLE hFile = CreateFile(sFullPathWorkingFile.c_str(),
        	GENERIC_WRITE, // Open for writing
            0, // No shared access
            NULL, // No security
            CREATE_ALWAYS, // Create new file and overwrite, if exists
            FILE_ATTRIBUTE_NORMAL,  // Normal file
            NULL); // No template

    	if (hFile == INVALID_HANDLE_VALUE) {
		 	_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
			sMessage.assign(L"Error creating file. Error code ").append(szHex);
			addMessageToEventList(sMessage.c_str());
			iRC = SI_FILEERROR;
		} else {
			DWORD dwBytesWritten = 0;
			// Write to CSV file
    		bool bErrorFlag = WriteFile(
                hFile, // File handle
                sData.c_str(), // Start of data to write
                sData.length()*sizeof(wchar_t), // Number of bytes to write
                &dwBytesWritten, // Number of bytes that were written
                NULL); // No overlapped structure

    		if (bErrorFlag == FALSE) {
			 	_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
				sMessage.assign(L"Error writing file. Error code ").append(szHex);
				addMessageToEventList(sMessage.c_str());
				iRC = SI_FILEERROR;
    		} else {
	    		if (dwBytesWritten != sData.length()*sizeof(wchar_t)) {
				 	_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
					sMessage.assign(L"Could write only ")
						.append(std::to_wstring(dwBytesWritten))
						.append(L" from ")
						.append(std::to_wstring(sData.length()*sizeof(wchar_t)))
						.append(L" bytes to file. Error code ").append(szHex);
					addMessageToEventList(sMessage.c_str());
					iRC = SI_FILEERROR;
				}
			}
			// Close CSV file
    		CloseHandle(hFile);
    	}
	}

	SendMessage(hWindow, WM_SERIALIMPORTFINISHED, iRC, 0); // Send window message to GUI

	sMessage.assign(L"Exit thread ID ").append(std::to_wstring(GetCurrentThreadId()));
	addMessageToEventList(sMessage.c_str());

	return iRC;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: importData

  Summary:  Reads from serial adapter and create CSV file

  Args:     HWND hWindow
              HWND of main window

  Returns:  BOOL
              TRUE = Success
			  FALSE = Error

-----------------------------------------------------------------F-F*/
BOOL importData(HWND hWindow) {
	std::wstring sCOMport = L"";
	std::wstring sMessage;

	if (hWindow == NULL) {
		MessageBox(NULL, L"Window handle is NULL",L"Error",MB_OK|MB_ICONERROR);
		return FALSE;
	}
	HWND hToolBar = GetDlgItem(hWindow, IDC_TOOLBAR);
    if (hToolBar == NULL) {
		addMessageToEventList(L"Error finding tool bar");
    	return FALSE;
	}
    if (g_hComboBox == NULL) {
		addMessageToEventList(L"Error finding combo box");
    	return FALSE;
	}

	if ((g_hThreadImportSerialData != NULL) && (WaitForSingleObject(g_hThreadImportSerialData,0) == WAIT_TIMEOUT)) {
		addMessageToEventList(L"Prevent duplicate serial import data");
		return FALSE;
	}

	if (SendMessage(g_hComboBox, (UINT) CB_GETCOUNT, (WPARAM) 0, (LPARAM) 0) == 0) {
		MessageBox(hWindow,LoadStringAsWstr(g_hInst,IDS_NOSERIALFOUND).c_str(), LoadStringAsWstr(g_hInst,IDS_MSGBOXINFOTITLE).c_str() ,MB_OK|MB_ICONINFORMATION);
		return FALSE;
	}
	DWORD dwSelection = SendMessage(g_hComboBox, (UINT) CB_GETCURSEL, (WPARAM) 0, (LPARAM) 0);
	if (dwSelection == CB_ERR) {
		MessageBox(hWindow,LoadStringAsWstr(g_hInst,IDS_NOSERIALSELECTED).c_str(), LoadStringAsWstr(g_hInst,IDS_MSGBOXINFOTITLE).c_str(),MB_OK|MB_ICONINFORMATION);
		return FALSE;
	}

	DWORD dwCOMport = SendMessage(g_hComboBox, (UINT) CB_GETITEMDATA, (WPARAM) dwSelection, (LPARAM) 0);
	if (dwCOMport == 0) {
		MessageBox(hWindow,LoadStringAsWstr(g_hInst,IDS_NOVALIDSERIALSELECTED).c_str(), LoadStringAsWstr(g_hInst,IDS_MSGBOXINFOTITLE).c_str(),MB_OK|MB_ICONINFORMATION);
		return FALSE;
	}

	// https://support.microsoft.com/en-us/topic/howto-specify-serial-ports-larger-than-com9-db9078a5-b7b6-bf00-240f-f749ebfd913e
	if (dwCOMport <= 9) {
		sCOMport = L"COM" + std::to_wstring(dwCOMport);
	} else {
		sCOMport = L"\\\\.\\COM" + std::to_wstring(dwCOMport);
	}

	addMessageToEventList(std::wstring(L"Opening '").append(sCOMport).append(L"").c_str());

	g_hSerialHandle = SerialInit(hWindow,sCOMport.c_str(),BAUDRATE);
	if (g_hSerialHandle == NULL) {
		sMessage.assign(LoadStringAsWstr(g_hInst,IDS_ERROROPEING)).append(L" '").append(sCOMport).append(L"'");
		if (GetLastError()== 0x0000001f) sMessage.append(L"\n").append(LoadStringAsWstr(g_hInst,IDS_FAKECH340WARNING));
		ErrorMessage(hWindow,sMessage,GetLastError());
		return FALSE;
	}
	if (g_hThreadImportSerialData != NULL) CloseHandle(g_hThreadImportSerialData);
	HWND* phWindow = (HWND *)malloc(sizeof(HWND));
	if (phWindow == NULL) {
		CloseHandle(g_hSerialHandle);
		g_hSerialHandle = NULL;
		sMessage.assign(LoadStringAsWstr(g_hInst,IDS_ERRORSERIALALLOCMEM)).append(L" '").append(sCOMport).append(L"'");
		ErrorMessage(hWindow,sMessage,0);
		return FALSE;
	}
	*phWindow = hWindow;
	g_hThreadImportSerialData = (HANDLE)_beginthreadex(0, 0, &threadImportSerialData, (void*)phWindow, 0, 0);

	return TRUE;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: setLastSelectedCOMfromSelection

  Summary:  Store selected COM adapter to to g_iLastSelectedCOM

  Args:

  Returns:
-----------------------------------------------------------------F-F*/
void setLastSelectedCOMfromSelection() {
	int iSelectedElement = SendMessage(g_hComboBox, (UINT) CB_GETCURSEL, (WPARAM) 0, (LPARAM) 0);
	if (iSelectedElement != CB_ERR) {
		g_iLastSelectedCOM = SendMessage(g_hComboBox, (UINT) CB_GETITEMDATA, (WPARAM) iSelectedElement, (LPARAM) 0);
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: addStaticCOMPorts

  Summary:  Add COM1-20 to combo box statically (for example under wine)

  Args:     HWND hWindow
              Handle to main window

  Returns:  unsigned int
  			  SA_OK = Success
  			  SA_ERROR = Error

-----------------------------------------------------------------F-F*/
unsigned int addStaticCOMPorts(HWND hWindow) {
	std::wstring sMessage;

    if (hWindow == NULL) { // Should not happen
		MessageBox(NULL, L"Window handle is NULL",L"Error",MB_OK|MB_ICONERROR);
    	return SA_ERROR;
	}

	sMessage.assign(L"Add static list of serial adapters");
	addMessageToEventList(sMessage.c_str());

	HWND hToolBar = GetDlgItem(hWindow, IDC_TOOLBAR);
    if (hToolBar == NULL)  {
		addMessageToEventList(L"Error finding tool bar");
		return SA_ERROR;
	}

    if (g_hComboBox == NULL) {
		addMessageToEventList(L"Error finding combo box");
		return SA_ERROR;
	}

	SendMessage(g_hComboBox, CB_RESETCONTENT, (WPARAM)0,(LPARAM) 0);
	#define MAXCOMPORTS 20
	for (int i=1;i<=MAXCOMPORTS;i++) {
		sMessage.assign(L"COM").append(std::to_wstring(i));
		int iElement = SendMessage(g_hComboBox,(UINT) CB_ADDSTRING,(WPARAM) 0,(LPARAM) sMessage.c_str());
		if ((iElement != CB_ERR) && (iElement != CB_ERRSPACE)) {
			SendMessage(g_hComboBox,(UINT) CB_SETITEMDATA,(WPARAM) iElement,(LPARAM) i);
			if (g_iLastSelectedCOM == i) { // Select last selected item
				SendMessage(g_hComboBox, CB_SETCURSEL, (WPARAM)iElement,(LPARAM) 0);
			}
		}
	}

	// Check selected item
    int iSelectedElement = SendMessage(g_hComboBox, CB_GETCURSEL, (WPARAM)0,(LPARAM) 0);
    if (iSelectedElement == CB_ERR) { // Nothing selected
   		SendMessage(g_hComboBox, CB_SETCURSEL, (WPARAM)0,(LPARAM) 0); // Select first serial adapter in list
		g_iLastSelectedCOM = SendMessage(g_hComboBox,(UINT) CB_GETITEMDATA,(WPARAM) 0,(LPARAM)0);
	}
	SendMessage(g_hComboBox, CB_SETCURSEL, (WPARAM)0,(LPARAM) 0); // Select first serial adapter in list
	setDisabledState(hWindow,IDM_IMPORT,false);
	return SA_OK;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: threadRescanSerialAdapter

  Summary:  Thread to detect serial adapters by WMI
            Based on https://learn.microsoft.com/en-us/windows/win32/wmisdk/example--getting-wmi-data-from-the-local-computer

  Args:     void *data
              Pointer to HWND of main window

  Returns:  unsigned int
  			  SA_OK = Success
  			  SA_ERROR = Error

-----------------------------------------------------------------F-F*/
unsigned int __stdcall threadRescanSerialAdapter(void *data) {
	HRESULT hres;
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT+2];
	std::wstring sMessage;

	HWND *phWindow = (HWND*) data;
	HWND hWindow = *phWindow;
	free(phWindow);

    if (hWindow == NULL) { // Should not happen
		MessageBox(NULL, L"Window handle is NULL",L"Error",MB_OK|MB_ICONERROR);
    	return SA_ERROR;
	}

	sMessage.assign(L"Rescan for serial adapters. Thread ID ").append(std::to_wstring(GetCurrentThreadId()));
	addMessageToEventList(sMessage.c_str());

	HWND hToolBar = GetDlgItem(hWindow, IDC_TOOLBAR);
    if (hToolBar == NULL)  {
		addMessageToEventList(L"Error finding tool bar");
		return SA_ERROR;
	}

    if (g_hComboBox == NULL) {
		addMessageToEventList(L"Error finding combo box");
		return SA_ERROR;
	}

	SendMessage(g_hComboBox, CB_RESETCONTENT, (WPARAM)0,(LPARAM) 0);

    // Step 1: --------------------------------------------------
    // Initialize COM. ------------------------------------------

    hres =  CoInitializeEx(0, COINIT_APARTMENTTHREADED); // COINIT_MULTITHREADED gives me a handle leak when CoInitializeEx was called just in threads
    if (FAILED(hres))
    {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", hres);
		sMessage.assign(L"Failed to initialize COM library. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
        return SA_ERROR; // Function has failed.
    }

    // Step 2: --------------------------------------------------
    // Set general COM security levels --------------------------

    hres =  CoInitializeSecurity(
        NULL,
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities
        NULL                         // Reserved
        );

    if (FAILED(hres) && (hres != RPC_E_TOO_LATE))
    {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", hres);
		sMessage.assign(L"Failed to initialize security. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
        CoUninitialize();
       	return SA_ERROR; // Function has failed.
    }

    // Step 3: ---------------------------------------------------
    // Obtain the initial locator to WMI -------------------------

    IWbemLocator *pLoc = NULL;

    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID *) &pLoc);

    if (FAILED(hres))
    {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", hres);
		sMessage.assign(L"Failed to create IWbemLocator object. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
        CoUninitialize();
        return SA_ERROR; // Function has failed.
    }

    // Step 4: -----------------------------------------------------
    // Connect to WMI through the IWbemLocator::ConnectServer method

    IWbemServices *pSvc = NULL;

    // Connect to the root\cimv2 namespace with
    // the current user and obtain pointer pSvc
    // to make IWbemServices calls.
    hres = pLoc->ConnectServer(
         _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
         NULL,                    // User name. NULL = current user
         NULL,                    // User password. NULL = current
         0,                       // Locale. NULL indicates current
         0,                    // Security flags.
         0,                       // Authority (for example, Kerberos)
         0,                       // Context object
         &pSvc                    // pointer to IWbemServices proxy
         );

    if (FAILED(hres))
    {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", hres);
		sMessage.assign(L"Could not connect. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
        pLoc->Release();
        CoUninitialize();
        return SA_ERROR; // Function has failed.
    }

	addMessageToEventList(L"Connected to ROOT\\CIMV2 WMI namespace");

    // Step 5: --------------------------------------------------
    // Set security levels on the proxy -------------------------

    hres = CoSetProxyBlanket(
       pSvc,                        // Indicates the proxy to set
       RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
       RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
       NULL,                        // Server principal name
       RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx
       RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
       NULL,                        // client identity
       EOAC_NONE                    // proxy capabilities
    );

    if (FAILED(hres))
    {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", hres);
		sMessage.assign(L"Could not set proxy blanket. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return SA_ERROR; // Function has failed.
    }

    // Step 6: --------------------------------------------------
    // Use the IWbemServices pointer to make requests of WMI ----

    // For example, get COM ports
    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_PnPEntity WHERE ClassGuid=\"{4d36e978-e325-11ce-bfc1-08002be10318}\""),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", hres);
		sMessage.assign(L"Query for COM ports failed. Error code ").append(szHex);
		addMessageToEventList(sMessage.c_str());

        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return SA_ERROR; // Function has failed.
    }

    // Step 7: -------------------------------------------------
    // Get the data from the query in step 6 -------------------

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;
   	int iCount = 0;
    do
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
            &pclsObj, &uReturn);

        if(uReturn != 0)
        {
	        VARIANT vtProp;

	        VariantInit(&vtProp);
	        // Get the value of the Name property
	        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);


		    std::wregex rx(L"\\(COM([0-9]+)\\)");
		    std::wcmatch mr;

		    if (std::regex_search(vtProp.bstrVal, mr, rx)) {
		        std::wcsub_match sub = mr[1]; // First sequence matched sub expression

				// Add items to combo box
				int iElement = SendMessage(g_hComboBox,(UINT) CB_ADDSTRING,(WPARAM) 0,(LPARAM) vtProp.bstrVal);
				int iCOM = _wtoi(sub.str().c_str());
				if ((iElement != CB_ERR) && (iElement != CB_ERRSPACE)) {
					SendMessage(g_hComboBox,(UINT) CB_SETITEMDATA,(WPARAM) iElement,(LPARAM)iCOM);
					if (g_iLastSelectedCOM == iCOM) { // Restore selection state of last selected item
   						SendMessage(g_hComboBox, CB_SETCURSEL, (WPARAM)iElement,(LPARAM) 0);
					}
				}
				iCount++;
		    }
	        VariantClear(&vtProp);

	        pclsObj->Release();
		};
    } while(uReturn != 0);

	sMessage.assign(L"Found ").append(std::to_wstring(iCount)).append(L" serial adapters");
	addMessageToEventList(sMessage.c_str());

    // Cleanup
    // ========

    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    CoUninitialize();

	// Check selected item, if combo box has at least one element
    if (iCount > 0) {
	    int iSelectedElement = SendMessage(g_hComboBox, CB_GETCURSEL, (WPARAM)0,(LPARAM) 0);
	    if (iSelectedElement == CB_ERR) { // Nothing selected
	   		SendMessage(g_hComboBox, CB_SETCURSEL, (WPARAM)0,(LPARAM) 0); // Select first serial adapter in list
			g_iLastSelectedCOM = SendMessage(g_hComboBox,(UINT) CB_GETITEMDATA,(WPARAM) 0,(LPARAM)0);
		}
		setDisabledState(hWindow,IDM_IMPORT,false);
	}
	sMessage.assign(L"Exit thread ID ").append(std::to_wstring(GetCurrentThreadId()));
	addMessageToEventList(sMessage.c_str());

	return SA_OK;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: rescanSerialAdapter

  Summary:  Start rescan for serial adapters

  Args:     HWND hWindow
              HWND of main window

  Returns:  BOOL
  			  TRUE = Success
  			  FALSE = Error

-----------------------------------------------------------------F-F*/
BOOL rescanSerialAdapter(HWND hWindow) {

	if ((g_hThreadRescanSerialAdapter != NULL) && (WaitForSingleObject(g_hThreadRescanSerialAdapter,0) == WAIT_TIMEOUT)) {
		addMessageToEventList(L"Prevent duplicate serial adapter rescan");
		return FALSE;
	}

	if (!isRunningUnderWine()) {

		if (g_hThreadRescanSerialAdapter != NULL) CloseHandle(g_hThreadRescanSerialAdapter);
		HWND* phWindow = (HWND *)malloc(sizeof(HWND));
		if (phWindow == NULL) {
			ErrorMessage(hWindow,LoadStringAsWstr(g_hInst,IDS_ERRORADAPTERALLOCMEM),0);
			return FALSE;
		}
		setDisabledState(hWindow,IDM_IMPORT,true);

		*phWindow = hWindow;
		g_hThreadRescanSerialAdapter = (HANDLE)_beginthreadex(0, 0, &threadRescanSerialAdapter, (void *)phWindow, 0, 0);
	} else { // Under wine we doe no have/use WMI => Fail back to static list COM1-20
		addMessageToEventList(L"Running under wine => static COM list");
		addStaticCOMPorts(hWindow);
	}

	return TRUE;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: copyCSV

  Summary:  Copy current CSV dataset file to a new CSV file

  Args:     HWND hWindow
              HWND of main window

  Returns:  BOOL
  			  TRUE = Success
  			  FALSE = Error

-----------------------------------------------------------------F-F*/
BOOL copyCSV(HWND hWindow) {
	OPENFILENAME ofn;

	wchar_t szWorkingFolder[MAX_PATH];
	GetModuleFileName(NULL,szWorkingFolder,MAX_PATH);
	PathRemoveFileSpec(szWorkingFolder);
	std::wstring sFullPathWorkingFile;
	sFullPathWorkingFile.assign(szWorkingFolder).append(L"\\").append(WORKINGCSVFILE);

	if (!PathFileExists(sFullPathWorkingFile.c_str())) {
		MessageBox(hWindow, LoadStringAsWstr(g_hInst,IDS_INFOIMPORTFIRST).c_str(),LoadStringAsWstr(g_hInst,IDS_MSGBOXINFOTITLE).c_str(),MB_OK|MB_ICONINFORMATION);
		return FALSE;
	}

    SYSTEMTIME tLocal;
    GetLocalTime(&tLocal);
	wchar_t szFileName[MAX_PATH];

	#define FILEPATTERN L"solarTemperatureLogger%04u%02u%02u%02d%02d%02d.csv"
	if (_snwprintf_s(szFileName, MAX_PATH, _TRUNCATE, FILEPATTERN, tLocal.wYear, tLocal.wMonth, tLocal.wDay, tLocal.wHour, tLocal.wMinute, tLocal.wSecond) < 0) {
		// Error
		return FALSE;
	}

	ZeroMemory(&ofn, sizeof(ofn));

	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = hWindow;
	std::wstring sResource(LoadStringAsWstr(g_hInst,IDS_FILESAVEFILTER));
	ofn.lpstrFilter = sResource.c_str();
	ofn.lpstrFile = szFileName;
	ofn.nMaxFile = MAX_PATH;
	ofn.lpstrDefExt = L"CSV";

	ofn.Flags = OFN_EXPLORER|OFN_PATHMUSTEXIST|OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT;
	if(GetSaveFileName(&ofn)) {
		// Copy file CSV file (Overwrites existing file)
		if(!CopyFile(sFullPathWorkingFile.c_str(), szFileName, FALSE)){
			wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT+2];
			_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
			std::wstring sMessage;
			sMessage.assign(L"Copy file failed '")
				.append(sFullPathWorkingFile)
				.append(L"' to '")
				.append(szFileName)
				.append(L"'. Error code ")
				.append(szHex);
			addMessageToEventList(sMessage.c_str());

			sMessage.assign(LoadStringAsWstr(g_hInst,IDS_ERRORCOPYFILE))
				.append(L" '")
				.append(sFullPathWorkingFile)
				.append(L"'")
				.append(LoadStringAsWstr(g_hInst,IDS_TO))
				.append(L"'")
				.append(szFileName)
				.append(L"'");
			ErrorMessage(hWindow,sMessage,GetLastError());
			return FALSE;
		}

		PathRemoveFileSpec(szFileName); // Remove filename from full path string to get the output folder
		openWithFileExplorer(hWindow,szFileName); // Open file explorer for the output folder
	}
	return TRUE;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: tableCompareFuncEx

  Summary:   Custom comparison function used for second column (needed because we have only strings in column >= 1)

  Args:     LPARAM lParam1
			  List view index of first item
  			LPARAM lParam2
  			  List view item of second item
			LPARAM lParamSort
			  Sort order and column
			  > 0               = ascending
			  <= 0              = descending
			  abs(lParamSort)-1 = column number

  Returns:  int
  			   0 = first and second item are equal
  			   1 = first item is greater than second item (when ascending is set)
  			  -1 = first item is less than second item (when ascending is set)

-----------------------------------------------------------------F-F*/
int CALLBACK tableCompareFuncEx(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	int iRC = 0;
	double dValue1, dValue2;
	#define MAXBUFFERSIZE 255
	wchar_t szBuffer1[MAXBUFFERSIZE+1];
	wchar_t szBuffer2[MAXBUFFERSIZE+1];

    BOOL isAsc = (lParamSort > 0);
    int nSortColumn = abs(lParamSort) -1;
	switch (nSortColumn) {
		case 1: // Sort second column
			ListView_GetItemText(g_hTable,lParam1,nSortColumn,szBuffer1,MAXBUFFERSIZE);
			dValue1 = _wtof(szBuffer1);
			ListView_GetItemText(g_hTable,lParam2,nSortColumn,szBuffer2,MAXBUFFERSIZE);
			dValue2 = _wtof(szBuffer2);
			if (dValue1 > dValue2) iRC=1;
			if (dValue1 < dValue2) iRC=-1;
			if (dValue1 == dValue2) iRC=0;
			break;
		default:
    		iRC = 0;
	}
    return isAsc ? iRC : -iRC;
}


/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: tableCompareFunc

  Summary:   Custom comparison function used for first column

  Args:     LPARAM lParam1
			  lParam value from column 0 for first item
  			LPARAM lParam2
			  lParam value from column 0 for for second item
			LPARAM lParamSort
			  Sort order and column
			  > 0     = ascending
			  <= 0    = descending

  Returns:  int
  			   0 = first item is less or equal than second item (when ascending is set)
  			   1 = first item is greater than second item (when ascending is set)

-----------------------------------------------------------------F-F*/
int CALLBACK tableCompareFunc(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
    BOOL isAsc = (lParamSort > 0);
    return isAsc ? lParam2 < lParam1 : lParam1 < lParam2;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: sortTableByColumn

  Summary:   Column header was clicked => Change sort order

  Args:     LPNMLISTVIEW pLVInfo
			  Pointer to the list view notification message
			  NULL = Reset sort to default

  Returns:

-----------------------------------------------------------------F-F*/
void sortTableByColumn(LPNMLISTVIEW pLVInfo)
{
    static int nSortColumn = 0;
    static BOOL bSortAscending = TRUE;
    LPARAM lParamSort;

	// Get table
	if (g_hTable == NULL) {
		addMessageToEventList(L"Table control not found");
		return;
	}

	if (pLVInfo != NULL) {

	    // Get new sort parameters
	    if (pLVInfo->iSubItem == nSortColumn) // When column has not change since last sorting
	        bSortAscending = !bSortAscending; // Toggle ascending/descending
	    else { // When column has changed
	        nSortColumn = pLVInfo->iSubItem;
	        bSortAscending = TRUE;
	    }

	} else { // Reset
		nSortColumn = 0;
		bSortAscending = TRUE;
	}

	std::wstring sLabel;
	LVCOLUMNW lvc;

	// Set text for first column header
	sLabel.assign(LoadStringAsWstr(g_hInst,IDS_COL0LABEL));
	if (nSortColumn == 0) {
		if (bSortAscending) sLabel.append(L" ▲"); else sLabel.append(L" ▼");
	}
   	lvc.mask = LVCF_TEXT;
   	lvc.pszText = const_cast<wchar_t *>(sLabel.c_str());
	ListView_SetColumn(g_hTable, 0, &lvc);

	// Set text for second column header
	sLabel.assign(LoadStringAsWstr(g_hInst,IDS_COL1LABEL));
	if (nSortColumn == 1) {
		if (bSortAscending) sLabel.append(L" ▲"); else sLabel.append(L" ▼");
	}
   	lvc.mask = LVCF_TEXT;
   	lvc.pszText = const_cast<wchar_t *>(sLabel.c_str());
	ListView_SetColumn(g_hTable, 1, &lvc);

    // Combine sort order and column into a single integer value
    lParamSort = 1 + nSortColumn; // Add column offset of 1 to prevent 0
    if (!bSortAscending)
        lParamSort = -lParamSort; // >0 = ascending, <0 = descending

    // Sort list
    switch (nSortColumn) {
    	case 1: // Second column
    		ListView_SortItemsEx(g_hTable, tableCompareFuncEx, lParamSort);
			break;
    	default: // First column
   		    ListView_SortItems(g_hTable, tableCompareFunc, lParamSort);
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: loadTableData

  Summary:  Load current dataset file to table view

  Args:     HWND hWindow
              HWND of main window

  Returns:  BOOL
  			  TRUE = Success
  			  FALSE = Error

-----------------------------------------------------------------F-F*/
BOOL loadTableData(HWND hWindow) {
	BOOL bValidCSVLine = TRUE;
	HANDLE hFile = INVALID_HANDLE_VALUE;
	BOOL bSuccess = FALSE;
	double minValue = 1000.0f;
	double maxValue = -273.15f;
	double sumValue = 0;
	// Locale ID to convert time string from CSV by VarDateFromStr
	LCID lcid = MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US), SORT_DEFAULT); // 0x409 is the locale ID for English US
	// String buffer for converting int to hex
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT+2];
	std::wstring sMessage;

	DATE dTime = 0;
    SYSTEMTIME stUTC {0}, stLocal {0};

	// Get table
	if (g_hTable == NULL) {
		addMessageToEventList(L"Table control not found");
		return FALSE;
	}

	// Get status bar
	HWND hStatusBar = GetDlgItem(hWindow, IDC_STATUSBAR);
	if (hStatusBar == NULL) {
		addMessageToEventList(L"Status bar control not found");
		return FALSE;
	}

	wchar_t szWorkingFolder[MAX_PATH];
	GetModuleFileName(NULL,szWorkingFolder,MAX_PATH);
	PathRemoveFileSpec(szWorkingFolder);
	std::wstring sFullPathWorkingFile;
	sFullPathWorkingFile.assign(szWorkingFolder).append(L"\\").append(WORKINGCSVFILE);

	addMessageToEventList(std::wstring(L"Read '")
		.append(sFullPathWorkingFile)
		.append(L"' with Locale ID ")
		.append(std::to_wstring(lcid))
		.c_str());

	// Read CSV file
	hFile = CreateFile(sFullPathWorkingFile.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, 0);
	if(hFile != INVALID_HANDLE_VALUE) {
		DWORD dwFileSize;
		dwFileSize = GetFileSize(hFile, NULL);
		if(dwFileSize != INVALID_FILE_SIZE) {
			LPWSTR pszFileText;
			pszFileText = (LPWSTR)GlobalAlloc(GPTR, (dwFileSize + 1)*sizeof(WCHAR));
			if(pszFileText != NULL) {
				DWORD dwRead;
				if(ReadFile(hFile, pszFileText, dwFileSize, &dwRead, NULL)) { // Read file
					if (dwFileSize != dwRead) { // File could not be loaded complete in buffer
						GlobalFree(pszFileText);
						CloseHandle(hFile);
						sMessage.assign(L"File '")
							.append(sFullPathWorkingFile)
							.append(L"' too big with ")
							.append(std::to_wstring(dwFileSize))
							.append(L" bytes");
						addMessageToEventList(sMessage.c_str());

						sMessage.assign(LoadStringAsWstr(g_hInst,IDS_FILE))
							.append(L" '")
							.append(sFullPathWorkingFile)
							.append(L"' ")
							.append(LoadStringAsWstr(g_hInst,IDS_WITH))
							.append(L" ")
							.append(std::to_wstring(dwFileSize))
							.append(L" ")
							.append(LoadStringAsWstr(g_hInst,IDS_BYTESTOGIG));
						ErrorMessage(hWindow, sMessage,0);
						return FALSE;
					}

					pszFileText[dwFileSize] = 0; // Set null terminator

					if (pszFileText[0] != 0xFEFF) { // Check unicode BOM (must bei UTF16 big endian)
						GlobalFree(pszFileText);
						CloseHandle(hFile);
						sMessage.assign(L"Unicode encoding in '")
							.append(sFullPathWorkingFile)
							.append(L"' is not UTF16 big endian");
						addMessageToEventList(sMessage.c_str());

						sMessage.assign(LoadStringAsWstr(g_hInst,IDS_FILE))
							.append(L" '")
							.append(sFullPathWorkingFile)
							.append(L"' ")
							.append(LoadStringAsWstr(g_hInst,IDS_ISNOT16BE));
						ErrorMessage(hWindow, sMessage.c_str(),0);
						return FALSE;
					};
					
					WaitForSingleObject(g_hSemaphoreDatasets,INFINITE); // Protect dataset vector

					// Clear datasets
					g_vDatasets.clear();
					
					// Clear table
					ListView_DeleteAllItems(g_hTable);
										
				    wchar_t* token = std::wcstok(pszFileText+1, L"\n") ; // Split string at newline

					int iLine = 0;
				    int iValidLines = 0;
				    int iNonValidLines = 0;
				    #define TIMESAMPLE L"01.01.2024 00:00:00"
				    while (token)
				    {
				    	std::wstring sLine(token);
				    	sLine.erase(sLine.find_last_not_of(L"\t\n\v\f\r ") + 1); // Right trim

				    	bValidCSVLine = TRUE;
				    	std::wstring::size_type seperatorPosition = sLine.find(L";");
				    	if (seperatorPosition == std::wstring::npos) bValidCSVLine = false; // Invalid line
				    	if (sLine.compare(L"UTC time;Degree celsius") == 0) {  // Header line
							addMessageToEventList(L"Valid CSV header detected");
							bValidCSVLine = false;
						}
						std::wstring sColumn1;
						std::wstring sColumn2;
						if (bValidCSVLine) {
  							sColumn1 = sLine.substr(0,seperatorPosition);
  							sColumn2 = sLine.substr(seperatorPosition+1);
  							if (sColumn1.length() != wcslen(TIMESAMPLE)) {
								addMessageToEventList(std::wstring(L"Time conversion to UTC failed. Token too short '")
									.append(sColumn1)
									.append(L"'")
									.append(L" in line ")
									.append(std::to_wstring(iLine))
									.c_str());
								bValidCSVLine = false;
							}
						}

						if (bValidCSVLine)  {
							// Convert UTC time string to local time, if possible
							int iRC = VarDateFromStr((wchar_t*)sColumn1.c_str(),lcid,0,&dTime);
							if (iRC != S_OK) {
								addMessageToEventList(std::wstring(L"Time conversion failed with return code ")
									.append(std::to_wstring(iRC))
									.append(L" for token '")
									.append(sColumn1).append(L"'")
									.append(L" in line ")
									.append(std::to_wstring(iLine))
									.c_str());
							} else {
								#define TIMEPATTERN L"%02u.%02u.%04u %02d:%02d:%02d"
						    	if (VariantTimeToSystemTime(dTime, &stUTC)) { // Get UTC time
						    		wchar_t buffer[sColumn1.length()+1];
						    		if (SystemTimeToTzSpecificLocalTime(NULL,&stUTC,&stLocal)) { // Get local time
										if (_snwprintf_s(buffer, sColumn1.length()+1, _TRUNCATE, TIMEPATTERN, stLocal.wDay, stLocal.wMonth, stLocal.wYear, stLocal.wHour, stLocal.wMinute, stLocal.wSecond) >= 0)
											sColumn1.assign(buffer);

									} else addMessageToEventList(std::wstring(L"Time conversion to local time failed for '")
										.append(sColumn1)
										.append(L"'")
										.append(L" in line ")
										.append(std::to_wstring(iLine))
										.c_str());

								} else {
									addMessageToEventList(std::wstring(L"Time conversion to UTC time failed for '")
										.append(token)
										.append(L"'")
										.append(L" in line ")
										.append(std::to_wstring(iLine))
										.c_str());
								}
							}

							// Add table row
 							LVITEMW lvi = {0};

				  			// First column with time
							lvi.mask = LVIF_TEXT | LVIF_PARAM;
							lvi.lParam = iValidLines;
							lvi.iItem = iValidLines+1; // Add at end of list
							lvi.pszText = (LPWSTR)sColumn1.c_str();
							ListView_InsertItem(g_hTable, &lvi);

							// Second column with data value
							// Replace , with .
						 	std::replace(sColumn2.begin(), sColumn2.end(), L',', L'.');
							ListView_SetItemText(g_hTable, iValidLines, 1, (LPWSTR)sColumn2.c_str());

							double value = _wtof(sColumn2.c_str());

							DATASET dataset {stUTC, value};
							g_vDatasets.push_back(dataset);
							sumValue += value;
							if (value > maxValue) maxValue = value;
							if (value < minValue) minValue = value;
							iValidLines++;
						}
						iLine ++;
						token = std::wcstok(nullptr, L"\n"); // Next line
				    }
				    sortTableByColumn(NULL); // Restore default sorting

					ReleaseSemaphore(g_hSemaphoreDatasets,1,NULL);

					InvalidateRect(g_hGraph,NULL,TRUE); // Force graph refresh

			    	addMessageToEventList(std::wstring(L"Found ")
						.append(std::to_wstring(iValidLines))
						.append(L" valid data sets in ")
						.append(std::to_wstring(iLine))
						.append(L" lines")
						.c_str());
					SendMessage(hStatusBar, SB_SETTEXT, 0, 
						(LPARAM)std::wstring(std::to_wstring(iValidLines))
						.append(L" ")
						.append(LoadStringAsWstr(g_hInst,IDS_DATASETS)).c_str());
					if (iValidLines > 1) {
						// Convert min/max to string and trim trailing zeros
						std::wstring sMinValue = std::wstring(std::to_wstring(minValue));
						std::wstring sMaxValue = std::wstring(std::to_wstring(maxValue));
						std::wstring sAvgValue = std::wstring(std::to_wstring(((int)std::round(10*sumValue/iValidLines))/10.0));
						for (std::wstring::iterator it=sMinValue.end()-1; it!=sMinValue.begin(); it--) if (*it == '0') *it = ' '; else break;
						for (std::wstring::iterator it=sMaxValue.end()-1; it!=sMaxValue.begin(); it--) if (*it == '0') *it = ' '; else break;
						for (std::wstring::iterator it=sAvgValue.end()-1; it!=sAvgValue.begin(); it--) if (*it == '0') *it = ' '; else break;
				    	sMinValue.erase(sMinValue.find_last_not_of(L" ") + 1);
				    	sMaxValue.erase(sMaxValue.find_last_not_of(L" ") + 1);
				    	sAvgValue.erase(sAvgValue.find_last_not_of(L" ") + 1);

						SendMessage(hStatusBar, SB_SETTEXT, 1, 
							(LPARAM)LoadStringAsWstr(g_hInst,IDS_MIN)
							.append(L": ")
							.append(sMinValue)
							.append(L" ")
							.append(LoadStringAsWstr(g_hInst,IDS_MAX))
							.append(L": ")
							.append(sMaxValue)
							.append(L" ")
							.append(LoadStringAsWstr(g_hInst,IDS_AVG))
							.append(L": ").append(sAvgValue).c_str());
					} else {
						SendMessage(hStatusBar, SB_SETTEXT, 1,(LPARAM)L"");
					}
					bSuccess = TRUE;
				}
				GlobalFree(pszFileText);
			}
		} else {
			_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
			sMessage.assign(L"Could not read file size for '")
				.append(sFullPathWorkingFile)
				.append(L"'. Error code ")
				.append(szHex);
			addMessageToEventList(sMessage.c_str());

			sMessage.assign(LoadStringAsWstr(g_hInst,IDS_ERRORGETFILESIZE))
				.append(L" '")
				.append(sFullPathWorkingFile)
				.append(L"'");
			ErrorMessage(hWindow,sMessage,GetLastError());
		}
		CloseHandle(hFile);
		setDisabledState(hWindow, IDM_SAVE, false);
	} else {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
		addMessageToEventList(std::wstring(L"Could not open file. Error code ")
			.append(szHex)
			.c_str());
		if (GetLastError() != 0x2) {
			
			sMessage.assign(LoadStringAsWstr(g_hInst,IDS_ERROROPEING))
				.append(L" '")
				.append(sFullPathWorkingFile)
				.append(L"'");
			ErrorMessage(hWindow,sMessage,GetLastError());
		}
	}

	return bSuccess;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: createMenu

  Summary:  Create menu items

  Args:     HWND hWindow
              HWND of main window

  Returns:

-----------------------------------------------------------------F-F*/
void createMenu(HWND hWindow) {
	// Main menu
	g_hMenu = CreateMenu();
	HMENU hFileMenu = CreateMenu();
	HMENU hExtraMenu = CreateMenu();

	AppendMenu(hFileMenu,MF_STRING | MF_GRAYED,IDM_IMPORT,LoadStringAsWstr(g_hInst,IDS_MENUIMPORT).c_str());
	AppendMenu(hFileMenu,MF_STRING | MF_GRAYED,IDM_SAVE,LoadStringAsWstr(g_hInst,IDS_MENUSAVE).c_str());
	AppendMenu(hFileMenu,MF_SEPARATOR,0,NULL);
	AppendMenu(hFileMenu,MF_STRING, IDM_EXPLOREFOLDER,LoadStringAsWstr(g_hInst,IDS_MENUOPENFOLDER).c_str());
	AppendMenu(hFileMenu,MF_STRING,IDM_INFO,LoadStringAsWstr(g_hInst,IDS_MENUINFO).c_str());
	AppendMenu(hFileMenu,MF_STRING,IDM_EXIT,LoadStringAsWstr(g_hInst,IDS_MENUEXIT).c_str());
	if (!isRunningUnderWine()) AppendMenu(hExtraMenu,MF_STRING,IDM_COMREFRESH,LoadStringAsWstr(g_hInst,IDS_MENUUPDATESERIAL).c_str());
	AppendMenu(hExtraMenu,MF_STRING,IDM_STATICCOM,LoadStringAsWstr(g_hInst,IDS_MENUSTATICSERIAL).c_str());
	if (!isRunningUnderWine()) AppendMenu(hExtraMenu,MF_STRING,IDM_DEVMGR,LoadStringAsWstr(g_hInst,IDS_MENUDEVICEMANAGER).c_str()); // No device manager in wine
	AppendMenu(g_hMenu,MF_POPUP,(UINT_PTR)hFileMenu,LoadStringAsWstr(g_hInst,IDS_MENUDATASETS).c_str());
	AppendMenu(g_hMenu,MF_POPUP,(UINT_PTR)hExtraMenu,LoadStringAsWstr(g_hInst,IDS_MENUEXTRAS).c_str());
	SetMenu(hWindow,g_hMenu);

	// Context menu
	g_hContextMenu = CreatePopupMenu();
	AppendMenu(g_hContextMenu,MF_STRING,IDM_COPYSELECTED,LoadStringAsWstr(g_hInst,IDS_MENUCOPYSELECTED).c_str());
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: logEnvironment

  Summary:  Add environmental data to event list

  Args:     HWND hWindow
              HWND of main window

  Returns:

-----------------------------------------------------------------F-F*/
void logEnvironment(HWND hWindow) {
	UINT uDpi = MyGetDpiForWindow(hWindow);

	wchar_t szExecutable[MAX_PATH];
	GetModuleFileName(NULL,szExecutable,MAX_PATH);
	addMessageToEventList(std::wstring(L"Executable=").append(szExecutable).c_str());
	addMessageToEventList(std::wstring(L"Monitors=").append(std::to_wstring(GetSystemMetrics(SM_CMONITORS))).c_str());
	addMessageToEventList(std::wstring(L"Virtual screen=").append(std::to_wstring(GetSystemMetrics(SM_CXVIRTUALSCREEN)))
		.append(L"x").append(std::to_wstring(GetSystemMetrics(SM_CYVIRTUALSCREEN))).c_str());
	addMessageToEventList(std::wstring(L"Primary screen=").append(std::to_wstring(GetSystemMetrics(SM_CXSCREEN)))
		.append(L"x").append(std::to_wstring(GetSystemMetrics(SM_CYSCREEN))).c_str());
	addMessageToEventList(std::wstring(L"DPI=").append(std::to_wstring(uDpi)).c_str());

	addMessageToEventList(std::wstring(L"Mousebuttons=").append(std::to_wstring(GetSystemMetrics(SM_CMOUSEBUTTONS))).c_str());
	addMessageToEventList(std::wstring(L"Multitouch=").append(std::to_wstring(GetSystemMetrics(SM_MAXIMUMTOUCHES))).c_str());

	addMessageToEventList(std::wstring(L"Bootmode=").append(std::to_wstring(GetSystemMetrics(SM_CLEANBOOT))).c_str());
	addMessageToEventList(std::wstring(L"Convertable mode=").append(std::to_wstring(!GetSystemMetrics(SM_CONVERTIBLESLATEMODE))).c_str());
	addMessageToEventList(std::wstring(L"Tablet PC=").append(std::to_wstring(GetSystemMetrics(SM_TABLETPC))).c_str());
	addMessageToEventList(std::wstring(L"Docked=").append(std::to_wstring(GetSystemMetrics(SM_SYSTEMDOCKED))).c_str());
	addMessageToEventList(std::wstring(L"Network=").append(std::to_wstring(GetSystemMetrics(SM_NETWORK)&1)).c_str());
	addMessageToEventList(std::wstring(L"Remote session=").append(std::to_wstring(GetSystemMetrics(SM_REMOTESESSION))).c_str());
	addMessageToEventList(std::wstring(L"Slow machine=").append(std::to_wstring(GetSystemMetrics(SM_SLOWMACHINE))).c_str());

	if (IsWindows10OrGreater()) addMessageToEventList(L"IsWindows10OrGreater");
	if (IsWindows8OrGreater()) addMessageToEventList(L"IsWindows8OrGreater");
	if (IsWindows8Point1OrGreater()) addMessageToEventList(L"IsWindows8Point1OrGreater");

	DWORD dwVersion = 0;
	DWORD dwMajorVersion = 0;
	DWORD dwMinorVersion = 0;
	DWORD dwBuild = 0;
	dwVersion = GetVersion();
	dwMajorVersion = (DWORD)(LOBYTE(LOWORD(dwVersion)));
	dwMinorVersion = (DWORD)(HIBYTE(LOWORD(dwVersion)));
	if (dwVersion < 0x80000000) dwBuild = (DWORD)(HIWORD(dwVersion));

	addMessageToEventList(std::wstring(L"Windows Version=")
		.append(std::to_wstring(dwMajorVersion))
		.append(L".")
		.append(std::to_wstring(dwMinorVersion))
		.append(L" (")
		.append(std::to_wstring(dwBuild))
		.append(L")")
		.c_str());

    NTSTATUS (NTAPI * RtlGetVersion)(PRTL_OSVERSIONINFOW lpVersionInformation) = nullptr;
    HMODULE hDLL = GetModuleHandle(L"ntdll.dll");
    if (hDLL == NULL)
        return;
    *reinterpret_cast<FARPROC*>(&RtlGetVersion) = GetProcAddress(hDLL, "RtlGetVersion");
    if (RtlGetVersion == nullptr)
        return;

    OSVERSIONINFOEX versionInfo{ sizeof(OSVERSIONINFOEX) };
    if (RtlGetVersion(reinterpret_cast<LPOSVERSIONINFO>(&versionInfo)) < 0) return;

	addMessageToEventList(std::wstring(L"Windows Real Version=")
		.append(std::to_wstring(versionInfo.dwMajorVersion))
		.append(L".")
		.append(std::to_wstring(versionInfo.dwMinorVersion))
		.append(L" (")
		.append(std::to_wstring(versionInfo.dwBuildNumber))
		.append(L")")
		.c_str());

	SYSTEM_INFO sysInfo;
    GetSystemInfo(&sysInfo);

    switch(sysInfo.wProcessorArchitecture)
    {
    case PROCESSOR_ARCHITECTURE_IA64:
  		addMessageToEventList(L"Architecture=Intel Itanium-based");
  		break;
    case PROCESSOR_ARCHITECTURE_ARM:
  		addMessageToEventList(L"Architecture=ARM");
  		break;
    case PROCESSOR_ARCHITECTURE_ARM64:
  		addMessageToEventList(L"Architecture=ARM64");
  		break;
    case PROCESSOR_ARCHITECTURE_AMD64:
  		addMessageToEventList(L"Architecture=x64 (AMD or Intel)");
  		break;
    case PROCESSOR_ARCHITECTURE_INTEL:
  		addMessageToEventList(L"Architecture=x86");
		break;
    case PROCESSOR_ARCHITECTURE_UNKNOWN:
  		addMessageToEventList(L"Architecture=unknown");
		break;
    }

    SYSTEM_INFO sysNativeInfo;
    GetNativeSystemInfo (&sysNativeInfo);

    switch(sysNativeInfo.wProcessorArchitecture)
    {
    case PROCESSOR_ARCHITECTURE_IA64:
  		addMessageToEventList(L"Native Architecture=Intel Itanium-based");
  		break;
    case PROCESSOR_ARCHITECTURE_ARM:
  		addMessageToEventList(L"Native Architecture=ARM");
  		break;
    case PROCESSOR_ARCHITECTURE_ARM64:
  		addMessageToEventList(L"Native Architecture=ARM64");
  		break;
    case PROCESSOR_ARCHITECTURE_AMD64:
  		addMessageToEventList(L"Native Architecture=x64 (AMD or Intel)");
  		break;
    case PROCESSOR_ARCHITECTURE_INTEL:
  		addMessageToEventList(L"Native Architecture=x86");
		break;
    case PROCESSOR_ARCHITECTURE_UNKNOWN:
  		addMessageToEventList(L"Native Architecture=unknown");
		break;

    }

}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: recreateToolBarButtons

  Summary:  The buttons on the toolbar does not scale when WM_DPICHANGED => Recreate buttons if needed

  Args:     HWND hWindow
              HWND of main window

  Returns:

-----------------------------------------------------------------F-F*/
void recreateToolBarButtons(HWND hWindow) {
	HWND hToolBar = GetDlgItem(hWindow, IDC_TOOLBAR);
	UINT uDpi = MyGetDpiForWindow(hWindow);

	if ((hToolBar != NULL)  && (g_hToolbarImageList16 != NULL) && (g_hToolbarImageList32 != NULL)) {
		// Delete all buttons
		for (int i=TOOLBARMAXITEMS-1;i>=0;i--) {
			SendMessage(hToolBar, TB_DELETEBUTTON, i, 0);
		}

	    TBBUTTON tbButtons[TOOLBARMAXITEMS] =
	    {
	        { 0, IDM_IMPORT, TBSTATE_ENABLED, BTNS_AUTOSIZE , {0}, 0, (INT_PTR)L"" },
	        { 1, IDM_SAVE, TBSTATE_ENABLED, BTNS_AUTOSIZE, {0}, 0, (INT_PTR)L"" },
	        { 2, IDM_INFO, TBSTATE_ENABLED, BTNS_AUTOSIZE , {0}, 0, (INT_PTR)L"" },
	        { 3, IDM_EXIT, TBSTATE_ENABLED, BTNS_AUTOSIZE, {0}, 0, (INT_PTR)L"" },
	        { 4, IDM_COMREFRESH, TBSTATE_ENABLED, BTNS_AUTOSIZE, {0}, 0, (INT_PTR)L"" },
	    };
	    SendMessage(hToolBar, TB_BUTTONSTRUCTSIZE, (WPARAM)sizeof(TBBUTTON), 0);

	    // Add buttons
	    SendMessage(hToolBar, TB_ADDBUTTONS, (WPARAM)TOOLBARMAXITEMS, (LPARAM)&tbButtons);

		// Restore state for import and save from menu
		UINT uSaveState = GetMenuState(g_hMenu,IDM_SAVE,MF_BYCOMMAND);
		if (uSaveState > 0) setDisabledState(hWindow, IDM_SAVE,(uSaveState & MF_DISABLED != 0));
		UINT uImportState = GetMenuState(g_hMenu,IDM_IMPORT,MF_BYCOMMAND);
		if (uImportState > 0) setDisabledState(hWindow, IDM_IMPORT,(uImportState & MF_DISABLED != 0));

	    // Resize the toolbar, and then show it
	    SendMessage(hToolBar, TB_AUTOSIZE, 0, 0);
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: alternatingLines

  Summary:   Create alternating background color for table lines

  Args:     HWND hDlg
  			  Handle to table

  Returns:

-----------------------------------------------------------------F-F*/
void alternatingLines(HWND hDlg)
{
    RECT rectUpdate, rectDest;
    COLORREF cBackgroundColor;

	// Check update area
    if (GetUpdateRect(hDlg, &rectUpdate, FALSE) == 0) return;

	// Number of all items
	int iAllItems = ListView_GetItemCount(hDlg);

    // "Paint" table as normal (Needed for table header, grids when table is empty...)
    if (iAllItems == 0) CallWindowProc(g_WindowProcTableOrig, hDlg, WM_PAINT,0,0);

    // First visible item
    int iFirstItem = ListView_GetTopIndex(hDlg);

    // Number for visible items
    int iVisibleItems = ListView_GetCountPerPage(hDlg);
    iVisibleItems ++; // ListView_GetCountPerPage counts only full visible items
    if (iFirstItem + iVisibleItems > iAllItems) iVisibleItems = iAllItems-iFirstItem;

   	RECT rectItem;
    for (int i = iFirstItem; i < iFirstItem + iVisibleItems; i++) {
    	ListView_GetItemRect(hDlg, i, &rectItem,LVIR_BOUNDS);
        if (IntersectRect(&rectDest, &rectUpdate, &rectItem)) {
        	// Set alternating background color
        	cBackgroundColor = GetSysColor(COLOR_WINDOW);
        	#define DARKENFACTOR 0.90f
        	if ((i % 2) == 0)  cBackgroundColor = RGB(
				(BYTE)(GetRValue(cBackgroundColor) * DARKENFACTOR),
        		(BYTE)(GetGValue(cBackgroundColor) * DARKENFACTOR),
        		(BYTE)(GetBValue(cBackgroundColor) * DARKENFACTOR));
            ListView_SetTextBkColor(hDlg, cBackgroundColor);

            // Invalidate line
        	InvalidateRect(hDlg, &rectItem, FALSE);

   			// Call original paint method
            CallWindowProc(g_WindowProcTableOrig, hDlg, WM_PAINT, 0, 0);
        }

    }
    // Set default background color
    cBackgroundColor = GetSysColor(COLOR_WINDOW);
    ListView_SetTextBkColor(hDlg, cBackgroundColor);    
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: WindowProcTable

  Summary:   Process window messages for the table (enables alternating background colors)

  Args:     HWND hDlg
              Handle to table
    		UINT uMsg
    		  Message
    		WPARAM wParam
    		LPARAM lParam

  Returns:  LRESULT

-----------------------------------------------------------------F-F*/
LRESULT CALLBACK WindowProcTable(
    HWND hDlg,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam)
{
	switch (uMsg) {
	    case WM_PAINT:
	        alternatingLines(hDlg);
	        return 0;
	}

	return CallWindowProc(g_WindowProcTableOrig,hDlg,uMsg,wParam,lParam);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: paintGraph

  Summary:   Paint simple data graph (should run flicker free)

  Args:     HWND hDlg
              Handle to graph area

  Returns:  

-----------------------------------------------------------------F-F*/
void paintGraph(HWND hDlg) {
    RECT  rect { 0,0,0,0 };
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hDlg, &ps);
    GetClientRect(hDlg, &rect);

	int iWidth = rect.right;
	int iHeight = rect.bottom;

	// Create memory device
	HDC hdcMemoryDevice = CreateCompatibleDC(hdc);
	if (hdcMemoryDevice == NULL) {
		addMessageToEventList(L"Error creating memory device");
	    EndPaint(hDlg, &ps);
		return;
	}
	// Create buffer bitmap
	HBITMAP bmBuffer = CreateCompatibleBitmap(hdc,iWidth, iHeight);
	if (bmBuffer == NULL) {
		addMessageToEventList(L"Error creating buffer bitmap");
		DeleteDC(hdcMemoryDevice);
	    EndPaint(hDlg, &ps);
		return;
	}
	// Backup memory device context state
	int iBackupDC = SaveDC(hdcMemoryDevice);
	
	HBITMAP hOldBitmap = (HBITMAP)SelectObject(hdcMemoryDevice,bmBuffer);
	
	FillRect(hdcMemoryDevice, &rect, (HBRUSH)(COLOR_WINDOW));

	WaitForSingleObject(g_hSemaphoreDatasets,INFINITE); // Protect dataset vector

	if (g_vDatasets.size() > 1) {

	 	double minValue = 1000.0f;
		double maxValue = -273.15f;
	
		for(const DATASET& dataset : g_vDatasets) {
			if (dataset.temperature > maxValue) maxValue = dataset.temperature;
			if (dataset.temperature < minValue) minValue = dataset.temperature;		
		}

		int deltaX = iWidth / (g_vDatasets.size()-1);
		int offsetX = (iWidth % (g_vDatasets.size()-1))/2;

		// Red selection marker(s)
		if (g_hTable != NULL) {
		    HPEN hPen = CreatePen(PS_SOLID, 3, RGB(255,0,0));
    		HPEN hOldPen = (HPEN)SelectObject(hdcMemoryDevice, hPen); 
			int iSelectItem = -1;
			while ((iSelectItem = ListView_GetNextItem(g_hTable, iSelectItem, LVNI_SELECTED)) > -1) {
				LV_ITEM Item {0};
   				Item.mask=LVIF_PARAM;
   				Item.iItem = iSelectItem;
   				Item.iSubItem = 0;
				if(ListView_GetItem(g_hTable, &Item)) {
		    		MoveToEx(hdcMemoryDevice, offsetX+Item.lParam*deltaX,0, NULL);
		    		LineTo(hdcMemoryDevice, offsetX+Item.lParam*deltaX,iHeight-1);	
				}
			}
    		SelectObject(hdcMemoryDevice, hOldPen);
    		DeleteObject(hPen);
		}

		// Black line for data values
		int X = 0;
		int Y = 0;
		for(const DATASET& dataset : g_vDatasets) {
			if ((maxValue-minValue) != 0.0f) {
				Y = (iHeight-1) - (dataset.temperature-minValue) * (iHeight-1) / (maxValue-minValue);
			} else {
				Y = iHeight/2; // Vertical centered line, if all values are equal
			}
			if (X == 0) { // First point
	    		MoveToEx(hdcMemoryDevice, offsetX+X,Y, NULL);
			} else { // Following points
	    		LineTo(hdcMemoryDevice, offsetX+X,Y);
			}
			X+=deltaX;	
		}

	}

	ReleaseSemaphore(g_hSemaphoreDatasets,1,NULL);
  
  	// Copy buffer to display
	BitBlt(hdc,0,0,iWidth,iHeight,hdcMemoryDevice,0,0,SRCCOPY);
	
	// Restore memory device context state
	SelectObject(hdcMemoryDevice,hOldBitmap);
	RestoreDC(hdcMemoryDevice,iBackupDC);
	DeleteObject(bmBuffer);
	DeleteDC(hdcMemoryDevice);

    EndPaint(hDlg, &ps);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: updateGraphSelection

  Summary:   Forces a dataset graph update, when a table row is selected

  Args:     

  Returns:  

-----------------------------------------------------------------F-F*/
void updateGraphSelection() {
	if (g_hTable == NULL) return; // Should not happen
	if (g_hGraph == NULL) return; // Should not happen
	
	int iSelectItem = ListView_GetNextItem(g_hTable, -1, LVNI_SELECTED);

	if (iSelectItem > -1) { // At lease one element is selected
		InvalidateRect(g_hGraph,NULL,TRUE); // Force graph update
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: WindowProcGraph

  Summary:   Process window messages for the graph area (replaces default paint and background erase)

  Args:     HWND hDlg
              Handle to table
    		UINT uMsg
    		  Message
    		WPARAM wParam
    		LPARAM lParam

  Returns:  LRESULT

-----------------------------------------------------------------F-F*/
LRESULT CALLBACK WindowProcGraph(
    HWND hDlg,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam)
{
	switch (uMsg) {
	    case WM_PAINT:
	    	paintGraph(hDlg); // Custom paint graph
	        return 1;
	    case WM_ERASEBKGND: // Background will be reset by paint
	    	return 1;
	}
	return CallWindowProc(g_WindowProcGraphOrig,hDlg,uMsg,wParam,lParam);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: createControls

  Summary:   Add all controls to window

  Args:     HWND hWindow
              Handle to main window

  Returns:

-----------------------------------------------------------------F-F*/
void createControls(HWND hWindow)
{
	UINT uDpi = MyGetDpiForWindow(hWindow);
    std::wstring sResource;

	// Remember font size for menu. Will be needed in resize function
	NONCLIENTMETRICS ncm;
    ncm.cbSize = sizeof(NONCLIENTMETRICS);
    SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
	g_iFontHeight_96DPI = MulDiv(ncm.lfMenuFont.lfHeight,USER_DEFAULT_SCREEN_DPI,uDpi);

    // Tool bar (With WS_EX_CONTROLPARENT WS_TABSTOP for the toolbar items does not work anymore, but for toolbar child items)
    HWND hToolBar = CreateWindowEx(
    	WS_EX_CONTROLPARENT,
		TOOLBARCLASSNAME,
		NULL,
        WS_CHILD | TBSTYLE_LIST | TBSTYLE_WRAPABLE | WS_TABSTOP | TBSTYLE_TOOLTIPS  ,
		0,0,0,0, // Position and size be set later
        hWindow,
		(HMENU)IDC_TOOLBAR,
		g_hInst,
		NULL);

	if(hToolBar !=NULL) {
	    recreateToolBarButtons(hWindow);
		ShowWindow(hToolBar, TRUE);
	}

	// Combo box
	g_hComboBox=CreateWindow(
		WC_COMBOBOX,
        NULL,
        CBS_DROPDOWN | CBS_HASSTRINGS | CBS_SIMPLE | WS_CHILD | WS_OVERLAPPED | WS_VISIBLE | WS_TABSTOP,
        0, 0, 0, 0, // Size will be set later
        hToolBar, // https://learn.microsoft.com/en-us/windows/win32/controls/embed-nonbutton-controls-in-toolbars (WS_TABSTOP does not work anymore for the combobox)
		(HMENU) IDC_COMBOBOX,
		g_hInst,
		(LPVOID) NULL);

   	// Create tooltip
    HWND hwndTooltip = CreateWindowEx(
		0,
		TOOLTIPS_CLASS,
		L"",	
		TTS_ALWAYSTIP, 
        0, 0, 0, 0, 
        hWindow, 0, g_hInst, 0);

    // Add tooltip for combobox
    TOOLINFO ti;
    ti.cbSize = sizeof(ti);    
    ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;     
    ti.uId = (UINT_PTR)g_hComboBox;    
	sResource.assign(LoadStringAsWstr(g_hInst,IDS_TOOLTIPSERIALADAPTER).c_str());
    ti.lpszText = const_cast<wchar_t *>(sResource.c_str());
    SendMessage(hwndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);

	// Status bar
	HWND hStatusBar = CreateWindow(
		STATUSCLASSNAME,
		NULL,
        SBARS_SIZEGRIP | WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, // Size will be set automatically
        hWindow,
		(HMENU)IDC_STATUSBAR,
		g_hInst,
		(LPVOID) NULL);

    int statusBarWidths[] = {0, -1}; // Size will be set later
    SendMessage(hStatusBar, SB_SETPARTS, sizeof(statusBarWidths)/sizeof(int), (LPARAM)statusBarWidths);

 	// Tabs
	HWND hTab = CreateWindow(
		WC_TABCONTROL,
		NULL,
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
		0,0,0,0, // Position and size be set later
        hWindow,
		(HMENU)IDC_TAB,
		g_hInst,
		(LPVOID) NULL);

	// Add tabs
	TCITEM tie;
	tie.mask = TCIF_TEXT;
	std::wstring sLabel;

	sLabel.assign(LoadStringAsWstr(g_hInst,IDS_DATASETS));
    tie.pszText = const_cast<wchar_t *>(sLabel.c_str());
    TabCtrl_InsertItem(hTab, 0, &tie);

	sLabel.assign(LoadStringAsWstr(g_hInst,IDS_LOG));
    tie.pszText = const_cast<wchar_t *>(sLabel.c_str());
    TabCtrl_InsertItem(hTab, 1, &tie);

	// List box
	g_hEventList = CreateWindow (
		L"LISTBOX",
        L"",
        WS_CHILD | WS_VSCROLL | WS_HSCROLL | LBS_NOTIFY | WS_TABSTOP | LBS_NOINTEGRALHEIGHT,
		0,0,0,0, // Position and size be set later
        hWindow, // To make WS_TABSTOP working this element is assigned to hWindow instead of hTab
		(HMENU) IDC_EVENTLIST,
		g_hInst,
		(LPVOID) NULL);

	// List view as table
	g_hTable = CreateWindow(
		WC_LISTVIEW,
		L"",
        WS_BORDER|WS_CHILD | LVS_REPORT | WS_VISIBLE | WS_TABSTOP,
		0,0,0,0, // Position and size be set later
        hWindow, // To make WS_TABSTOP working this element is assigned to hWindow instead of hTab
		(HMENU)IDC_TABLE,
		g_hInst,
		(LPVOID) NULL);
	ListView_SetExtendedListViewStyle(g_hTable, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES );
	// Change window procedure to enable custom paint table...
	g_WindowProcTableOrig = (WNDPROC)SetWindowLongPtr(g_hTable,GWLP_WNDPROC,(LONG_PTR)WindowProcTable);

	// Fill table
	LVCOLUMN lvc;
	
	// Table header
	lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
	lvc.fmt = LVCFMT_LEFT;
	lvc.cx = MulDiv(INITIALTABLELEFTWIDTH_96DPI, uDpi, USER_DEFAULT_SCREEN_DPI);
	sResource.assign(LoadStringAsWstr(g_hInst,IDS_COL0LABEL).c_str());
	lvc.pszText = const_cast<wchar_t *>(sResource.c_str());
	lvc.iSubItem = 0;
	ListView_InsertColumn(g_hTable, 0, & lvc); // First column

	lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
	lvc.fmt = LVCFMT_LEFT;
	lvc.cx = 0; // Will be set later
	sResource.assign(LoadStringAsWstr(g_hInst,IDS_COL1LABEL).c_str());
	lvc.pszText = const_cast<wchar_t *>(sResource.c_str());
	lvc.iSubItem = 1;
	ListView_InsertColumn(g_hTable, 1, & lvc); // Second column

	// Area for simple data graph
	g_hGraph = CreateWindowW(
		L"static",
		L"",
		WS_VISIBLE | WS_CHILD,
		0,0,0,0, // Position and size be set later
		hWindow,
		(HMENU) IDC_GRAPH,
		g_hInst,
		(LPVOID) NULL);
	// Change window procedure to enable custom paint for graph...
	g_WindowProcGraphOrig = (WNDPROC)SetWindowLongPtr(g_hGraph,GWLP_WNDPROC,(LONG_PTR)WindowProcGraph);

	// Static text for progress bar
	HWND hWndStaticText = CreateWindowW(
		L"static",
		L"",
		WS_CHILD|SS_CENTERIMAGE|SS_RIGHT,
		0,0,0,0, // Position and size be set later
		hWindow,
		(HMENU) IDC_PROGRESSTEXT,
		g_hInst,
		(LPVOID) NULL);

	// Progress bar
	HWND hProgessBar = CreateWindow(
		PROGRESS_CLASS,
		(LPTSTR) NULL,
        WS_CHILD | PBS_SMOOTH,
		0,0,0,0, // Position and size be set later
        hWindow,
		(HMENU) IDC_PROGRESSBAR,
		g_hInst,
		(LPVOID) NULL);

	// Button for progress bar
	HWND hButtonAbort = CreateWindow(
	    L"BUTTON",
	    LoadStringAsWstr(g_hInst,IDS_ABORT).c_str(),
	    WS_TABSTOP | WS_CHILD,
		0,0,0,0, // Position and size be set later
	    hWindow,
	    (HMENU) IDC_PROGRESSBUTTON,
	    g_hInst,
	    (LPVOID) NULL);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: copySelectedEventToClipboard

  Summary:   Copy selected line in event list to clipboard

  Args:

  Returns:  bool
  			  true = Success
  			  false = Error

-----------------------------------------------------------------F-F*/
bool copySelectedEventToClipboard() {
	if (g_hEventList == NULL) return false;

    int count = SendMessage(g_hEventList, LB_GETCOUNT, 0, 0);
    int iSelected = 0;
    for (int i = 0; i < count; i++)
    {
    	iSelected = SendMessage(g_hEventList, LB_GETSEL, i, 0);
    	if (iSelected == 1) {
           	std::wstring str;
           	int length = SendMessage(g_hEventList, LB_GETTEXTLEN, (WPARAM) i, 0);

        	wchar_t szLine[length+1];
            SendMessage(g_hEventList, LB_GETTEXT, i, (LPARAM)szLine);

            if (!OpenClipboard(NULL)) return false;
        	HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE | GMEM_DDESHARE, (wcslen(szLine) + 1)*sizeof(WCHAR));
        	memcpy(GlobalLock(hMem), szLine, (wcslen(szLine) +1) *sizeof(WCHAR));

			GlobalUnlock(hMem);
			EmptyClipboard();
			SetClipboardData(CF_UNICODETEXT, hMem);
			CloseClipboard();
		}
    }
    return true;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: overwriteMinMaxInfo

  Summary:   Set minimum window size when WM_GETMINMAXINFO message was received

  Args:     HWND hWindow
              Handle to main window
    		LPMINMAXINFO lpMMI
    		  Pointer to window minimum/maximum structure

  Returns:

-----------------------------------------------------------------F-F*/
void overwriteMinMaxInfo(HWND hWindow, LPMINMAXINFO lpMMI) {
	UINT uDpi = MyGetDpiForWindow(hWindow);
	
	// Desired client size
	RECT rectDesiredSize;
	rectDesiredSize.left = 0;
	rectDesiredSize.top = 0;
	rectDesiredSize.right = MulDiv(WINDOWMINWIDTH_96DPI, uDpi, USER_DEFAULT_SCREEN_DPI);
	rectDesiredSize.bottom = MulDiv(WINDOWMINHEIGHT_96DPI, uDpi, USER_DEFAULT_SCREEN_DPI);
	// Get window size for desired client size
	AdjustWindowRectEx(&rectDesiredSize,GetWindowLong(hWindow,GWL_STYLE),FALSE,GetWindowLong(hWindow,GWL_EXSTYLE));

	// Set window minimum size
	lpMMI->ptMinTrackSize.x = rectDesiredSize.right - rectDesiredSize.left;
	lpMMI->ptMinTrackSize.y = rectDesiredSize.bottom - rectDesiredSize.top;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: resizeControls

  Summary:   Resize/reposition window controls

  Args:     HWND hWindow
              Handle to main window

  Returns:

-----------------------------------------------------------------F-F*/
void resizeControls(HWND hWindow) {
	LONG units = GetDialogBaseUnits();
	UINT uDpi = MyGetDpiForWindow(hWindow);

    if (hWindow == NULL) return;

	// Get handle to controls
	HWND hStatusBar = GetDlgItem(hWindow, IDC_STATUSBAR);
    HWND hTab = GetDlgItem(hWindow, IDC_TAB);
	HWND hToolBar = GetDlgItem(hWindow, IDC_TOOLBAR);
	HWND hProgessBar = GetDlgItem(hWindow, IDC_PROGRESSBAR);
	HWND hProgressButton = GetDlgItem(hWindow, IDC_PROGRESSBUTTON);
	HWND hProgressText = GetDlgItem(hWindow, IDC_PROGRESSTEXT);

    // Check handles
    if (hStatusBar == NULL) return;
    if (hTab == NULL) return;
    if (hToolBar == NULL) return;
	if (hProgessBar == NULL) return;
	if (hProgressButton == NULL) return;
	if (hProgressText == NULL) return;
	if (g_hComboBox == NULL) return;
    if (g_hTable == NULL) return;
    if (g_hEventList == NULL) return;
    if (g_hGraph == NULL) return;

	// Get height/width of window inner area
    RECT rectWindow {0,0,0,0};
    GetClientRect(hWindow, &rectWindow);
	int iClientWidth = rectWindow.right;
	int iClientHeight = rectWindow.bottom;

	// Set icons for taskbar and caption
	SendMessage(hWindow, WM_SETICON, ICON_BIG, (LONG)(ULONG_PTR)LoadImage(g_hInst,MAKEINTRESOURCE(IDI_ICON32),IMAGE_ICON,32,32,LR_DEFAULTCOLOR | LR_SHARED) );
	if (uDpi > 144) {
		SendMessage(hWindow, WM_SETICON, ICON_SMALL, (LONG)(ULONG_PTR)LoadImage(g_hInst,MAKEINTRESOURCE(IDI_ICON32),IMAGE_ICON,32,32,LR_DEFAULTCOLOR | LR_SHARED) );
	} else {
		SendMessage(hWindow, WM_SETICON, ICON_SMALL, (LONG)(ULONG_PTR)LoadImage(g_hInst,MAKEINTRESOURCE(IDI_ICON16),IMAGE_ICON,16,16,LR_DEFAULTCOLOR | LR_SHARED) );
	}

	// Change toolbar image list dependent on current DPI
	if (uDpi >= 144)
		SendMessage(hToolBar, TB_SETIMAGELIST, (WPARAM)0,  (LPARAM)g_hToolbarImageList32);
	else
		SendMessage(hToolBar, TB_SETIMAGELIST, (WPARAM)0,  (LPARAM)g_hToolbarImageList16);
	SendMessage(hToolBar, TB_AUTOSIZE , (WPARAM)0, (LPARAM)0);

	// Update automatic size of status bar
    if (hStatusBar != NULL ) SendMessage(hStatusBar, WM_SIZE, 0, 0);

	// Sets font size for tabs and list from menu font size
	NONCLIENTMETRICS ncm;
    ncm.cbSize = sizeof(NONCLIENTMETRICS);
    SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);

    ncm.lfMenuFont.lfHeight = MulDiv(g_iFontHeight_96DPI, uDpi, USER_DEFAULT_SCREEN_DPI);

	if (g_hFont != NULL) DeleteObject(g_hFont); // Prevents GDI leak
    g_hFont = CreateFontIndirect(&ncm.lfMenuFont);
    SendMessage(g_hEventList, WM_SETFONT, (WPARAM) g_hFont, MAKELPARAM(TRUE, 0));
    SendMessage(hStatusBar, WM_SETFONT, (WPARAM) g_hFont, MAKELPARAM(TRUE, 0));
    SendMessage(hTab, WM_SETFONT, (WPARAM) g_hFont, MAKELPARAM(TRUE, 0));
    SendMessage(g_hTable, WM_SETFONT, (WPARAM) g_hFont, MAKELPARAM(TRUE, 0));
    SendMessage(hToolBar, WM_SETFONT, (WPARAM) g_hFont, MAKELPARAM(TRUE, 0));
    SendMessage(g_hComboBox, WM_SETFONT, (WPARAM) g_hFont, MAKELPARAM(TRUE, 0));
    SendMessage(hProgressButton, WM_SETFONT, (WPARAM) g_hFont, MAKELPARAM(TRUE, 0));
    SendMessage(hProgressText, WM_SETFONT, (WPARAM) g_hFont, MAKELPARAM(TRUE, 0));

	// Get toolbar size
    RECT rectToolBar {0,0,0,0};
	GetClientRect(hToolBar, &rectToolBar);
	int iToolBarWidth = rectToolBar.right;
	int iToolBarHeight = rectToolBar.bottom;

	// Get status bar size
    RECT rectStatusBar {0,0,0,0};
	GetClientRect(hStatusBar, &rectStatusBar);
	int iStatusBarWidth = rectStatusBar.right;
	int iStatusBarHeight = rectStatusBar.bottom;

	// Resize combobox
    RECT rectComboBox {0,0,0,0};
   	GetWindowRect(g_hComboBox, &rectComboBox);
   	int iComboBoxHeight = (rectComboBox.bottom-rectComboBox.top)+1;
	DWORD dwPadding = 0;
	dwPadding = SendMessage(hToolBar, TB_GETPADDING , (WPARAM) 0, (LPARAM)0);
	RECT rectButton {0,0,0,0};
	SendMessage(hToolBar, TB_GETITEMRECT , (WPARAM) (TOOLBARMAXITEMS-1), (LPARAM)&rectButton); // Get rect of last button
	int iButtonHeight = (rectButton.bottom - rectButton.top)+1;
	// Center combobox right to last button
	SetWindowPos(g_hComboBox, 0,rectButton.right+1, (2*rectButton.top+iButtonHeight-iComboBoxHeight)/2,
		iToolBarWidth-(rectButton.right+1)-LOWORD(dwPadding),
		iComboBoxHeight, SWP_NOZORDER);

	// Set width of left and right side of statusbar
    int statwidths[] = {MulDiv(STATUSBARLEFTWIDTH_96DPI, uDpi, USER_DEFAULT_SCREEN_DPI) , -1};
    SendMessage(hStatusBar, SB_SETPARTS, sizeof(statwidths)/sizeof(int), (LPARAM)statwidths);

	// Resize tab
	SetWindowPos(hTab, 0, 0,  iToolBarHeight, iClientWidth, iClientHeight-iToolBarHeight-iStatusBarHeight, SWP_NOZORDER);

	// Get tab inner area
	RECT rectTabInner{0,0,0,0};
    GetClientRect(hTab, &rectTabInner);
    TabCtrl_AdjustRect(hTab, FALSE, &rectTabInner);

	// Event list
	SetWindowPos(g_hEventList, 0, 
		rectTabInner.left,
		iToolBarHeight+ rectTabInner.top,
		(rectTabInner.right-rectTabInner.left)+1,
		(rectTabInner.bottom-rectTabInner.top)+1, SWP_NOZORDER);
	// Resize listbox width to ~double window width
	SendMessage(g_hEventList, LB_SETHORIZONTALEXTENT, ((rectTabInner.right-rectTabInner.left)+1)*2, 0 );

	int iGraphHeight = MulDiv(GRAPHHEIGHT_96DPI, uDpi, USER_DEFAULT_SCREEN_DPI);
	// Table
	SetWindowPos(g_hTable, 0, 
		rectTabInner.left,
		iToolBarHeight+rectTabInner.top,
		(rectTabInner.right-rectTabInner.left)+1,
		(rectTabInner.bottom-rectTabInner.top)+1-iGraphHeight, SWP_NOZORDER);
	// Graph
	SetWindowPos(g_hGraph, 0, 
		rectTabInner.left,
		iToolBarHeight+rectTabInner.bottom-iGraphHeight+1,
		(rectTabInner.right-rectTabInner.left)+1,
		iGraphHeight, SWP_NOZORDER);

	 // Expand second column
	LVCOLUMN lvc;
	lvc.mask = LVCF_WIDTH;
	ListView_GetColumn(g_hTable, 0, &lvc);
	int iAllItems = ListView_GetItemCount(g_hTable);
    int iVisibleItems = ListView_GetCountPerPage(g_hTable);
    if (iVisibleItems < iAllItems) {
		lvc.cx = (rectTabInner.right-rectTabInner.left)+1-lvc.cx-MyGetSystemMetricsForDpi(SM_CXVSCROLL, uDpi); // With scrollbar
	} else {
		lvc.cx = (rectTabInner.right-rectTabInner.left)+1-lvc.cx; // Without scrollbar
	}

	ListView_SetColumn(g_hTable, 1, &lvc);

	// Text left of progress bar
	int iButtonWidth = MulDiv(HIWORD(units), 50, 4);
	SetWindowPos(hProgressText, 0, 0,
		iClientHeight - iStatusBarHeight,
		iStatusBarWidth-MulDiv(STATUSBARLEFTWIDTH_96DPI, uDpi, USER_DEFAULT_SCREEN_DPI)-iButtonWidth-MulDiv(HIWORD(units), 3, 4),
		iStatusBarHeight, SWP_NOZORDER);

	// Progress bar
	SetWindowPos(hProgessBar, 0, iStatusBarWidth-MulDiv(STATUSBARLEFTWIDTH_96DPI, uDpi, USER_DEFAULT_SCREEN_DPI)-iButtonWidth,
		iClientHeight - iStatusBarHeight,
		MulDiv(STATUSBARLEFTWIDTH_96DPI, uDpi, USER_DEFAULT_SCREEN_DPI),
		iStatusBarHeight, SWP_NOZORDER);

	// Abort button right of progress bar
	SetWindowPos(hProgressButton, 0, iStatusBarWidth-iButtonWidth,
		iClientHeight - iStatusBarHeight,
		iButtonWidth,
		iStatusBarHeight, SWP_NOZORDER);

  	InvalidateRect(hWindow, NULL, TRUE); // Force window content update
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: sendAbortSignal

  Summary:   Send abort signal to import thread

  Args:     HWND hWindow
              Handle to main window

  Returns:  bool
  				true = success
  				false = failed

-----------------------------------------------------------------F-F*/
bool sendAbortSignal(HWND hWindow)
{
	HWND hProgressButton = GetDlgItem(hWindow, IDC_PROGRESSBUTTON);
	HWND hProgressText = GetDlgItem(hWindow, IDC_PROGRESSTEXT);

	if ((hProgressButton == NULL) || (hProgressText == NULL)) {
		addMessageToEventList(L"Error finding controls");
    	return false;
	}

	SetEvent(g_hImportAbortEvent);
	EnableWindow(hProgressButton,false);
	SendMessage(hProgressText, WM_SETTEXT, (WPARAM)0, (LPARAM)(LoadStringAsWstr(g_hInst,IDS_ABORTPENDING).c_str()));
	return true;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: openContextMenu

  Summary:   Context menu for event list

  Args:     HWND hWindow
              Handle to main window
            HWND hControl
              Handle to window control, where context menu was requested for
            int X
            	Mouse X position
            int Y
            	Mouse Y position

  Returns:  bool
  				true = success
  				false = failed

-----------------------------------------------------------------F-F*/
bool openContextMenu(HWND hWindow, HWND hControl, int X, int Y) {
	if (hWindow == NULL) { // Should not happen
		MessageBox(NULL, L"Window handle is NULL",L"Error",MB_OK|MB_ICONERROR);
		return false;
	}

	if (g_hEventList == NULL) return false; // Should not happen

	if (!IsWindowVisible(g_hEventList)) return false; // Context menu only, if event list is not hidden

    HWND hTab = GetDlgItem(hWindow, IDC_TAB);
	// WC_TABCONTROL catches all WM_CONTEXTMENU message for his clients => Search for tab client
	if (hControl == hTab) {
		POINT position = { X, Y };
		ScreenToClient(hControl, &position);
		hControl= ChildWindowFromPoint(hControl,position);
	}

	if (g_hEventList != hControl) return false; // Currently context menu only for the event list

    int iCount = SendMessage(g_hEventList, LB_GETCOUNT, 0, 0);
	if ((iCount) <= 0) return false; // Context menu only, if event list is not empty

	if ((X==0xffff) || (Y == 0xffff)) { // Shift+F10 or App button
		RECT rectItem = { 0 };
    	for (int i = 0; i < iCount; i++) { // Search for selected item
	    	if (SendMessage(g_hEventList, LB_GETSEL, i, 0) == 1) {
	    		if (SendMessage(g_hEventList,LB_GETITEMRECT,i, (LPARAM) &rectItem) != LB_ERR) {
	    			// Get X/Y position of selected item
   					POINT position = { rectItem.left, rectItem.top };
					ClientToScreen(g_hEventList, &position);
					X = position.x;
					Y = position.y;
				}
	    	}
		}
    }
	if ((X==0xffff) || (Y == 0xffff)) return false; // No item

	// Get and select item on X/Y position
	POINT position = { X, Y };
	ScreenToClient(g_hEventList, &position);
    int iItem = SendMessage(g_hEventList, LB_ITEMFROMPOINT, 0, MAKELPARAM(position.x,position.y));
    SendMessage(g_hEventList, LB_SETCURSEL , LOWORD(iItem), 0);

    // Open menu
	TrackPopupMenuEx(g_hContextMenu, TPM_LEFTALIGN + TPM_TOPALIGN, X, Y, hWindow, NULL);
	return true;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: showProgramInformation

  Summary:   Show program information message box

  Args:     HWND hWindow
              Handle to main window

  Returns:  

-----------------------------------------------------------------F-F*/
void showProgramInformation(HWND hWindow) {
	std::wstring sTitle(LoadStringAsWstr(g_hInst,IDS_WINDOWTITLE));
	wchar_t szExecutable[MAX_PATH];
	GetModuleFileName(NULL,szExecutable,MAX_PATH);

	DWORD  verHandle = 0;
	UINT   size      = 0;
	LPBYTE lpBuffer  = NULL;
	DWORD  verSize   = GetFileVersionInfoSize( szExecutable, NULL);
	
	if (verSize != 0) {
	    BYTE *verData = new BYTE[verSize];
	
	    if (GetFileVersionInfo( szExecutable, verHandle, verSize, verData))
	    {
	        if (VerQueryValue(verData,L"\\",(VOID FAR* FAR*)&lpBuffer,&size)) {
	            if (size) {
	                VS_FIXEDFILEINFO *verInfo = (VS_FIXEDFILEINFO *)lpBuffer;
	                if (verInfo->dwSignature == 0xfeef04bd) {
						sTitle.append(L" ")
							.append(std::to_wstring((verInfo->dwFileVersionMS >> 16 ) & 0xffff))
							.append(L".")
							.append(std::to_wstring((verInfo->dwFileVersionMS >> 0 ) & 0xffff))
							.append(L".")
							.append(std::to_wstring((verInfo->dwFileVersionLS >> 16 ) & 0xffff))
							.append(L".")
							.append(std::to_wstring((verInfo->dwFileVersionLS >> 0 ) & 0xffff));
	                }
	            }
	        }
	    }
	    delete[] verData;
	}

	MessageBox(hWindow,LoadStringAsWstr(g_hInst,IDS_PROGINFO).c_str(),sTitle.c_str(),MB_ICONINFORMATION|MB_OK);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: WindowProc

  Summary:   Process window messages of main window

  Args:     HWND hDlg
              Handle to main window
    		UINT uMsg
    		  Message
    		WPARAM wParam
    		LPARAM lParam

  Returns:  LRESULT

-----------------------------------------------------------------F-F*/
LRESULT CALLBACK WindowProc(
    HWND hWindow,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam)
{
   	switch (uMsg)
    {
    	case WM_SERIALIMPORTFINISHED:
			loadTableData(hWindow);
			resizeControls(hWindow);
    		switch (wParam) {
				case SI_TIMEOUT:
					MessageBox(hWindow,LoadStringAsWstr(g_hInst,IDS_IMPORTTIMEOUT).c_str(),LoadStringAsWstr(g_hInst,IDS_MSGBOXWARNINGTITLE).c_str(),MB_ICONWARNING|MB_OK);
					break;
				case SI_COMERROR:
					ErrorMessage(hWindow,LoadStringAsWstr(g_hInst,IDS_IMPORTSERIALERROR),0);
					break;
				case SI_FILEERROR:
					ErrorMessage(hWindow,LoadStringAsWstr(g_hInst,IDS_IMPORTFILEERROR),0);
					break;
				case SI_ABORT:
					MessageBox(hWindow,LoadStringAsWstr(g_hInst,IDS_IMPORTABORTED).c_str(),LoadStringAsWstr(g_hInst,IDS_MSGBOXINFOTITLE).c_str(),MB_ICONINFORMATION|MB_OK);
					break;
			}
    		break;
	    case WM_SIZE:
	        resizeControls(hWindow);
	    	break;
	    case WM_CLOSE:
	        DestroyWindow(hWindow);
	        break;
	    case WM_DESTROY:
	        PostQuitMessage(0);
	        break;
		case WM_CREATE:
			createMenu(hWindow);
			createControls(hWindow);
			loadTableData(hWindow);
			rescanSerialAdapter(hWindow);
			logEnvironment(hWindow);
			break;
		case WM_DEVICECHANGE:
			switch (wParam) {
				case DBT_DEVICEREMOVECOMPLETE:
					addMessageToEventList(L"Device removal");
					rescanSerialAdapter(hWindow);
					break;
				case DBT_DEVICEARRIVAL:
					addMessageToEventList(L"Device arrival");
					rescanSerialAdapter(hWindow);
					break;
			}
			break;
		case WM_DPICHANGED: {
			addMessageToEventList(std::wstring(L"DPI changed to ").append(std::to_wstring(MyGetDpiForWindow(hWindow))).c_str());
			RECT* const prcNewWindow = (RECT*)lParam;
			// Resize window
        	if (prcNewWindow != NULL) SetWindowPos(hWindow,
	            NULL,
	            prcNewWindow ->left,
	            prcNewWindow ->top,
	            prcNewWindow->right - prcNewWindow->left,
	            prcNewWindow->bottom - prcNewWindow->top,
	            SWP_NOZORDER | SWP_NOACTIVATE);
	        recreateToolBarButtons(hWindow);
	        resizeControls(hWindow);
			break;
		}
		case WM_GETMINMAXINFO:
			overwriteMinMaxInfo(hWindow,(LPMINMAXINFO)lParam);
			break;
		case WM_NOTIFY:
		    switch (((LPNMHDR)lParam)->code)
		    {
    			case TTN_GETDISPINFO:
		        {
		        	// Tooltips for toolbar
		            LPTOOLTIPTEXT lpttt = (LPTOOLTIPTEXT)lParam;

		            // Set the instance of the module that contains the resource.
		            lpttt->hinst = g_hInst;

		            UINT_PTR idButton = lpttt->hdr.idFrom;

		            switch (idButton) {
			            case IDM_IMPORT:
							_snwprintf_s(lpttt->szText, 100, _TRUNCATE, L"%s", LoadStringAsWstr(g_hInst,IDS_TOOLTIPIMPORT).c_str());
			                break;
			            case IDM_SAVE:
							_snwprintf_s(lpttt->szText, 100, _TRUNCATE, L"%s", LoadStringAsWstr(g_hInst,IDS_TOOLTIPSAVE).c_str());
			                break;
			            case IDM_INFO:
							_snwprintf_s(lpttt->szText, 100, _TRUNCATE, L"%s", LoadStringAsWstr(g_hInst,IDS_TOOLTIPPROGINFO).c_str());
			                break;
			            case IDM_EXIT:
							_snwprintf_s(lpttt->szText, 100, _TRUNCATE, L"%s", LoadStringAsWstr(g_hInst,IDS_TOOLTIPEXIT).c_str());
			                break;
			            case IDM_COMREFRESH:
							_snwprintf_s(lpttt->szText, 100, _TRUNCATE, L"%s", LoadStringAsWstr(g_hInst,IDS_TOOLTIPUPDATESERIAL).c_str());
			                break;
		        	}
		            break;
		        	}
		    	case LVN_COLUMNCLICK:
		        	sortTableByColumn((LPNMLISTVIEW)lParam);
		            break;
	    		case LVN_ITEMCHANGED:
	    			updateGraphSelection();
	    			break;
		   		case TCN_SELCHANGE:
		   			switch(TabCtrl_GetCurSel(GetDlgItem(hWindow, IDC_TAB))) {
		   				case 0:
			   				ShowWindow(g_hTable,true);
			   				ShowWindow(g_hGraph,true);
			   				ShowWindow(g_hEventList,false);
		   					break;
		   				case 1:
			   				ShowWindow(g_hEventList,true);
			   				ShowWindow(g_hTable,false);
			   				ShowWindow(g_hGraph,false);
		   					break;
					}
	            	break;
		    }
	    	break;
	    case WM_COMMAND:
	    	switch (HIWORD(wParam)) {
		    	case LBN_SELCHANGE: // same value as CBN_SELCHANGE
					switch (LOWORD(wParam)) {
			  			case IDC_COMBOBOX:
			  				setLastSelectedCOMfromSelection();
							break;
					}
					break;
		    	case BN_CLICKED:
		    		switch (LOWORD(wParam)) {
		    			case IDM_IMPORT:
		    				importData(hWindow);
		    				break;
		    			case IDM_SAVE:
		    				copyCSV(hWindow);
		    				break;
		    			case IDM_EXIT:
		    				DestroyWindow(hWindow);
		    				break;
		    			case IDM_COMREFRESH:
							rescanSerialAdapter(hWindow);
		    				break;
		    			case IDM_STATICCOM:
		    				addStaticCOMPorts(hWindow);
		    				break;
		    			case IDM_INFO:
		    				showProgramInformation(hWindow);
		    				break;
						case IDC_PROGRESSBUTTON:
							sendAbortSignal(hWindow);
							break;
						case IDM_COPYSELECTED:
							copySelectedEventToClipboard();
							break;
						case IDM_EXPLOREFOLDER:
							openWithFileExplorer(hWindow,L".");
							break;
						case IDM_DEVMGR:
							// Open device manager as admin
							if (((INT_PTR) ShellExecute(hWindow, L"runAs", L"devmgmt.msc", NULL, NULL, SW_SHOWNORMAL)) == SE_ERR_ACCESSDENIED) {
								// Retry als normal user, if access was denied
								ShellExecute(hWindow, L"open", L"devmgmt.msc", NULL, NULL, SW_SHOWNORMAL);
							}
							break;
					}
		 			break;
		 	}
		case WM_CONTEXTMENU:
			openContextMenu(hWindow,(HWND)wParam,LOWORD(lParam), HIWORD(lParam));
			break;
	    default:
	        return DefWindowProc(hWindow, uMsg, wParam, lParam);
    }
    return 0;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: WinMain

  Summary:   Process window messages of main window

  Args:     HINSTANCE hInstance
  			HINSTANCE hPrevInstance
			LPSTR lpszCmdLine
			int nCmdShow

  Returns:  int
  			  0 = success
  			  1 = error

-----------------------------------------------------------------F-F*/
int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
    LPSTR lpszCmdLine, int nCmdShow)
{
    MSG msg;
    BOOL bRet;
    WNDCLASS wc;
    UNREFERENCED_PARAMETER(lpszCmdLine);
	std::wstring sMessage = L"";
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT+2];


	// Create a semaphore for dataset vector
    g_hSemaphoreDatasets = CreateSemaphore( 
        NULL,           // default security attributes
        1,  // initial count
        1,  // maximum count
        NULL);          // unnamed semaphore
    if (g_hSemaphoreDatasets == NULL) {
   		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
    	sMessage.assign(L"CreateSemaphore failed. Error code ").append(szHex);
		MessageBox(NULL,sMessage.c_str(),L"Error!",MB_ICONERROR|MB_OK|MB_SYSTEMMODAL);
        return 1;
    }

    // Register the window class for the main window.
    if (!hPrevInstance)
    {
        wc.style = 0;
        wc.lpfnWndProc = (WNDPROC) WindowProc;
        wc.cbClsExtra = 0;
        wc.cbWndExtra = 0;
        wc.hInstance = hInstance;
        wc.hIcon =  (HICON) LoadImage(g_hInst, MAKEINTRESOURCE(IDI_ICON16), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR);
        wc.hCursor = LoadCursor((HINSTANCE) NULL,  IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW);
        wc.lpszMenuName =  L"MainMenu";
        wc.lpszClassName = L"MainWndClass";

        if (!RegisterClass(&wc))
        {
        	_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
     	    sMessage.assign(L"RegisterClass failed. Error code ").append(szHex);
			MessageBox(NULL,sMessage.c_str(),L"Error!",MB_ICONERROR|MB_OK|MB_SYSTEMMODAL);
            return 1;
   		}
    }

    g_hInst = hInstance;  // Save instance handle

 	// Enables controls from Comctl32.dll, like status bar, tabs ...
  	InitCommonControls();

	// Create image lists for the tool bar
	g_hToolbarImageList32 = ImageList_Create(32, 32, ILC_COLOR16 | ILC_MASK,  TOOLBARMAXITEMS, 0); // 32x32
	for (int i=0;i<TOOLBARMAXITEMS;i++) {
    	HBITMAP bitmap =  (HBITMAP)LoadImage(g_hInst, MAKEINTRESOURCE(IDB_IMPORT32+i),IMAGE_BITMAP,32,32,0);
		ImageList_AddMasked(g_hToolbarImageList32, bitmap, CLR_DEFAULT);
		DeleteObject(bitmap);
	}

	g_hToolbarImageList16 = ImageList_Create(16, 16, ILC_COLOR16 | ILC_MASK,  TOOLBARMAXITEMS, 0); // 16x16
	for (int i=0;i<TOOLBARMAXITEMS;i++) {
    	HBITMAP bitmap =  (HBITMAP)LoadImage(g_hInst, MAKEINTRESOURCE(IDB_IMPORT16+i),IMAGE_BITMAP,16,16,0);
		ImageList_AddMasked(g_hToolbarImageList16, bitmap, CLR_DEFAULT);
		DeleteObject(bitmap);
	}

   	g_hImportAbortEvent = CreateEvent(
        NULL,               // default security attributes
        FALSE,               // manual-reset event
        FALSE,              // initial state is nonsignaled
        TEXT("ExitEvent")  // object name
        );

    if (g_hImportAbortEvent == NULL)
    {
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
  	    sMessage.assign(L"CreateEvent failed. Error code ").append(szHex);
		MessageBox(NULL,sMessage.c_str(),L"Error!",MB_ICONERROR|MB_OK|MB_SYSTEMMODAL);
		return 1;
    }

    // Create the main window.
    HWND hwndMain = CreateWindow(
		L"MainWndClass",
		LoadStringAsWstr(g_hInst,IDS_WINDOWTITLE).c_str(),
        WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
        CW_USEDEFAULT,
		CW_USEDEFAULT,
		(HWND) NULL,
        (HMENU) NULL,
		g_hInst,
		(LPVOID) NULL);

    // If the main window cannot be created, terminate the application.
    if (hwndMain==NULL)
    {
  		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
  	    sMessage.assign(L"CreateWindow failed. Error code ").append(szHex);
		MessageBox(NULL,sMessage.c_str(),L"Error!",MB_ICONERROR|MB_OK|MB_SYSTEMMODAL);
		return 1;
	}

    // Show the window and paint its contents
    ShowWindow(hwndMain, nCmdShow);
    UpdateWindow(hwndMain);

    // Start the message loop
    while( (bRet = GetMessage( &msg, NULL, 0, 0 )) != 0)
    {
        if (bRet == -1)
        {
 			_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT+2, _TRUNCATE, L"0x%08X", GetLastError());
  	    	sMessage.assign(L"GetMessage failed. Error code ").append(szHex);
			MessageBox(NULL,sMessage.c_str(),L"Error!",MB_ICONERROR|MB_OK|MB_SYSTEMMODAL);
			return 1;
        }
        else
        {
        	if (!IsDialogMessage(hwndMain, &msg))
			{
            	TranslateMessage(&msg);
            	DispatchMessage(&msg);
			}
        }
    }

    // Cleanup
    DeleteObject(g_hFont);
    ImageList_Destroy(g_hToolbarImageList16);
    ImageList_Destroy(g_hToolbarImageList32);

	CloseHandle(g_hThreadRescanSerialAdapter);
	TerminateThread(g_hThreadImportSerialData,0);
	CloseHandle(g_hThreadImportSerialData);
	CloseHandle(g_hSerialHandle);
	CloseHandle(g_hImportAbortEvent);
	CloseHandle(g_hSemaphoreDatasets);

    // Return the exit code to the system.
    return msg.wParam;
}
