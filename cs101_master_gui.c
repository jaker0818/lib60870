/*
 * cs101_master_gui.c
 * IEC 60870-5-101 Master with GUI Interface
 */

#define UNICODE
#define _UNICODE

#include <windows.h>
#include <commctrl.h>
#include <setupapi.h>
#include <devguid.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "hal_time.h"
#include "hal_thread.h"
#include "hal_serial.h"
#include "cs101_master.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "lib60870.lib")
#pragma comment(lib, "setupapi.lib")

#define ID_COMBO_PORT 100
#define ID_COMBO_BAUDRATE 101
#define ID_COMBO_DATABITS 102
#define ID_COMBO_PARITY 103
#define ID_COMBO_STOPBITS 104
#define ID_EDIT_SLAVE_ADDR 105
#define ID_EDIT_MASTER_ADDR 106
#define ID_BTN_CONNECT 107
#define ID_BTN_DISCONNECT 108
#define ID_BTN_INTERROGATION 109
#define ID_BTN_SEND_CMD 110
#define ID_BTN_TIME_SYNC 111
#define ID_EDIT_IOA 112
#define ID_EDIT_CMD_VALUE 113
#define ID_EDIT_LOG 114
#define ID_BTN_CLEAR_LOG 115
#define ID_BTN_REFRESH 116
#define ID_STATUS_TEXT 117

// Global variables
HINSTANCE hInst;
HWND hWndMain;
HWND hLogEdit;
HWND hStatusText;

SerialPort port = NULL;
CS101_Master master = NULL;
bool isConnected = false;
bool running = false;
char selectedPort[20] = "COM1";
int baudRate = 9600;
int dataBits = 8;
char parity = 'E';
int stopBits = 1;
int slaveAddr = 1;
int masterAddr = 2;
DWORD threadId = NULL;
HANDLE hWorkerThread = NULL;
CRITICAL_SECTION logCS;

// Add text to log window with timestamp
void AddLog(const char* text, BOOL isSend)
{
    SYSTEMTIME st;
    char buffer[1024];
    char timestamp[32];
    wchar_t wbuffer[1024];

    GetLocalTime(&st);
    sprintf(timestamp, "[%02d:%02d:%02d.%03d] ", st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);

    EnterCriticalSection(&logCS);

    int len = GetWindowTextLengthW(hLogEdit);
    SendMessageW(hLogEdit, EM_SETSEL, len, len);
    sprintf(buffer, "\r\n%s%s %s", timestamp, isSend ? "SEND:" : "RECV:", text);
    MultiByteToWideChar(CP_UTF8, 0, buffer, -1, wbuffer, 1024);
    SendMessageW(hLogEdit, EM_REPLACESEL, FALSE, (LPARAM)wbuffer);

    // Auto-scroll
    SendMessageW(hLogEdit, EM_SETSEL, -1, -1);
    SendMessageW(hLogEdit, EM_SCROLLCARET, 0, 0);

    LeaveCriticalSection(&logCS);
}

// Raw message handler
static void rawMessageHandler(void* parameter, uint8_t* msg, int msgSize, bool sent)
{
    char buffer[512];
    int pos = 0;
    int i;

    for (i = 0; i < msgSize && pos < sizeof(buffer) - 4; i++) {
        pos += sprintf(buffer + pos, "%02X ", msg[i]);
    }
    buffer[pos] = '\0';

    AddLog(buffer, sent);
}

// ASDU received handler
static bool asduReceivedHandler(void* parameter, int address, CS101_ASDU asdu)
{
    char buffer[512];
    sprintf(buffer, "ASDU: %s(%i) CA:%d Elements:%d",
            TypeID_toString(CS101_ASDU_getTypeID(asdu)),
            CS101_ASDU_getTypeID(asdu),
            address,
            CS101_ASDU_getNumberOfElements(asdu));
    AddLog(buffer, FALSE);

    if (CS101_ASDU_getTypeID(asdu) == M_ME_TE_1) {
        int i;
        for (i = 0; i < CS101_ASDU_getNumberOfElements(asdu); i++) {
            MeasuredValueScaledWithCP56Time2a io =
                (MeasuredValueScaledWithCP56Time2a) CS101_ASDU_getElement(asdu, i);

            if (io) {
                sprintf(buffer, "  IOA:%d Value:%d",
                        InformationObject_getObjectAddress((InformationObject) io),
                        MeasuredValueScaled_getValue((MeasuredValueScaled) io));
                AddLog(buffer, FALSE);
                MeasuredValueScaledWithCP56Time2a_destroy(io);
            }
        }
    }
    else if (CS101_ASDU_getTypeID(asdu) == M_SP_NA_1) {
        int i;
        for (i = 0; i < CS101_ASDU_getNumberOfElements(asdu); i++) {
            SinglePointInformation io =
                (SinglePointInformation) CS101_ASDU_getElement(asdu, i);

            if (io) {
                sprintf(buffer, "  IOA:%d Value:%d",
                        InformationObject_getObjectAddress((InformationObject) io),
                        SinglePointInformation_getValue((SinglePointInformation) io));
                AddLog(buffer, FALSE);
                SinglePointInformation_destroy(io);
            }
        }
    }

    return true;
}

// Link layer state handler
static void linkLayerStateChanged(void* parameter, int address, LinkLayerState state)
{
    const char* stateStr = "";
    char buffer[256];

    switch (state) {
        case LL_STATE_IDLE:
            stateStr = "IDLE";
            break;
        case LL_STATE_ERROR:
            stateStr = "ERROR";
            break;
        case LL_STATE_BUSY:
            stateStr = "BUSY";
            break;
        case LL_STATE_AVAILABLE:
            stateStr = "AVAILABLE";
            break;
    }

    sprintf(buffer, "Link Layer: %s", stateStr);
    wchar_t wbuffer[256];
    MultiByteToWideChar(CP_UTF8, 0, buffer, -1, wbuffer, 256);
    SetWindowTextW(hStatusText, wbuffer);
}

// Worker thread for master operation
DWORD WINAPI MasterWorkerThread(LPVOID lpParam)
{
    int cycleCounter = 0;

    while (running && isConnected) {
        if (master) {
            CS101_Master_run(master);
        }

        Thread_sleep(10);
        cycleCounter++;
    }

    return 0;
}

// 扫描串口 - 使用注册表方法（支持虚拟串口）
void ScanSerialPortsFromRegistry(HWND hCombo)
{
    HKEY hKey;
    DWORD index = 0;
    char portName[256];
    DWORD portNameSize = sizeof(portName);
    char friendlyName[256];
    DWORD friendlyNameSize = sizeof(friendlyName);
    char displayName[512];
    int portCount = 0;

    SendMessageW(hCombo, CB_RESETCONTENT, 0, 0);

    // 打开注册表键
    if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, "HARDWARE\\DEVICEMAP\\SERIALCOMM", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        while (RegEnumValueA(hKey, index++, portName, &portNameSize, NULL, NULL, (LPBYTE)friendlyName, &friendlyNameSize) == ERROR_SUCCESS) {
            // 转换为宽字符
            wchar_t wport[256];
            MultiByteToWideChar(CP_UTF8, 0, portName, -1, wport, 256);

            // 添加到列表
            SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)wport);
            portCount++;

            // 重置大小
            portNameSize = sizeof(portName);
            friendlyNameSize = sizeof(friendlyName);
        }
        RegCloseKey(hKey);
    }

    // 如果注册表方法没有找到串口，使用遍历方法
    if (portCount == 0) {
        AddLog("No serial ports found in registry, scanning COM1-COM256...", FALSE);
        for (int i = 1; i <= 256; i++) {
            char portName[20];
            sprintf(portName, "COM%d", i);
            
            wchar_t wport[20];
            MultiByteToWideChar(CP_UTF8, 0, portName, -1, wport, 20);
            
            SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)wport);
        }
    } else {
        char buffer[256];
        sprintf(buffer, "Found %d serial ports (including virtual ports)", portCount);
        AddLog(buffer, FALSE);
    }

    // 默认选择第一个串口
    if (SendMessageW(hCombo, CB_GETCOUNT, 0, 0) > 0) {
        SendMessageW(hCombo, CB_SETCURSEL, 0, 0);
    }
}

// 扫描串口 - 使用 SetupAPI 方法（更详细的信息）
void ScanSerialPortsFromSetupAPI(HWND hCombo)
{
    HDEVINFO deviceInfoSet;
    SP_DEVINFO_DATA deviceInfoData;
    DWORD index = 0;
    int portCount = 0;
    char buffer[512];

    SendMessageW(hCombo, CB_RESETCONTENT, 0, 0);

    // 创建设备信息集
    deviceInfoSet = SetupDiGetClassDevsA(&GUID_DEVCLASS_PORTS, NULL, NULL, DIGCF_PRESENT);

    if (deviceInfoSet == INVALID_HANDLE_VALUE) {
        AddLog("Failed to get device info set, using registry method...", FALSE);
        ScanSerialPortsFromRegistry(hCombo);
        return;
    }

    deviceInfoData.cbSize = sizeof(SP_DEVINFO_DATA);

    // 枚举所有串口设备
    while (SetupDiEnumDeviceInfo(deviceInfoSet, index, &deviceInfoData)) {
        HKEY hKey;
        char portName[256] = "";
        DWORD size = sizeof(portName);
        BOOL found = FALSE;

        // 打开设备的注册表键
        hKey = SetupDiOpenDevRegKey(deviceInfoSet, &deviceInfoData, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ);

        if (hKey != INVALID_HANDLE_VALUE) {
            // 读取端口名
            if (RegQueryValueExA(hKey, "PortName", NULL, NULL, (LPBYTE)portName, &size) == ERROR_SUCCESS) {
                found = TRUE;
            }
            RegCloseKey(hKey);
        }

        // 如果找到了串口，添加到列表
        if (found && strncmp(portName, "COM", 3) == 0) {
            wchar_t wport[256];
            MultiByteToWideChar(CP_UTF8, 0, portName, -1, wport, 256);
            
            SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)wport);
            portCount++;

            sprintf(buffer, "Found port: %s", portName);
            AddLog(buffer, FALSE);
        }

        index++;
    }

    SetupDiDestroyDeviceInfoList(deviceInfoSet);

    // 如果没有找到串口，使用注册表方法作为备用
    if (portCount == 0) {
        AddLog("No ports found via SetupAPI, trying registry...", FALSE);
        ScanSerialPortsFromRegistry(hCombo);
    } else {
        sprintf(buffer, "Total serial ports found: %d (including virtual ports)", portCount);
        AddLog(buffer, FALSE);

        // 默认选择第一个串口
        if (SendMessageW(hCombo, CB_GETCOUNT, 0, 0) > 0) {
            SendMessageW(hCombo, CB_SETCURSEL, 0, 0);
        }
    }
}

// Initialize serial ports list - 扫描可用串口
void InitializePortCombo(HWND hCombo)
{
    AddLog("Scanning for available serial ports...", FALSE);
    ScanSerialPortsFromSetupAPI(hCombo);
}

// Initialize baud rate list
void InitializeBaudrateCombo(HWND hCombo)
{
    const char* rates[] = {"300", "600", "1200", "2400", "4800", "9600", "19200", "38400", "57600", "115200"};
    for (int i = 0; i < 10; i++) {
        wchar_t wrate[20];
        MultiByteToWideChar(CP_UTF8, 0, rates[i], -1, wrate, 20);
        SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)wrate);
    }
    SendMessageW(hCombo, CB_SETCURSEL, 5, 0);
}

// Initialize data bits list
void InitializeDataBitsCombo(HWND hCombo)
{
    const char* bits[] = {"5", "6", "7", "8"};
    for (int i = 0; i < 4; i++) {
        wchar_t wbit[20];
        MultiByteToWideChar(CP_UTF8, 0, bits[i], -1, wbit, 20);
        SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)wbit);
    }
    SendMessageW(hCombo, CB_SETCURSEL, 3, 0);
}

// Initialize parity list
void InitializeParityCombo(HWND hCombo)
{
    const char* parities[] = {"None (N)", "Odd (O)", "Even (E)", "Mark (M)", "Space (S)"};
    for (int i = 0; i < 5; i++) {
        wchar_t wparity[20];
        MultiByteToWideChar(CP_UTF8, 0, parities[i], -1, wparity, 20);
        SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)wparity);
    }
    SendMessageW(hCombo, CB_SETCURSEL, 2, 0);
}

// Initialize stop bits list
void InitializeStopBitsCombo(HWND hCombo)
{
    const char* bits[] = {"1", "1.5", "2"};
    for (int i = 0; i < 3; i++) {
        wchar_t wbit[20];
        MultiByteToWideChar(CP_UTF8, 0, bits[i], -1, wbit, 20);
        SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)wbit);
    }
    SendMessageW(hCombo, CB_SETCURSEL, 0, 0);
}

// Refresh serial ports list
void OnRefreshPorts()
{
    HWND hCombo = GetDlgItem(hWndMain, ID_COMBO_PORT);
    
    AddLog("Refreshing serial port list...", FALSE);
    InitializePortCombo(hCombo);
    
    MessageBoxW(hWndMain, L"串口列表已刷新", L"提示", MB_ICONINFORMATION);
}

// Connect to slave
void OnConnect()
{
    char buffer[256];

    if (isConnected) {
        MessageBoxW(hWndMain, L"Already connected!", L"Error", MB_ICONERROR);
        return;
    }

    // Get port name
    HWND hPort = GetDlgItem(hWndMain, ID_COMBO_PORT);
    int portIndex = SendMessageW(hPort, CB_GETCURSEL, 0, 0);
    if (portIndex == CB_ERR) {
        MessageBoxW(hWndMain, L"请选择串口!", L"Error", MB_ICONERROR);
        return;
    }
    
    wchar_t wport[20];
    SendMessageW(hPort, CB_GETLBTEXT, portIndex, (LPARAM)wport);
    WideCharToMultiByte(CP_UTF8, 0, wport, -1, selectedPort, 20, NULL, NULL);

    // 对于COM10及以上的串口，Windows需要 "\\\\.\\COMXX" 格式
    char portName[100];
    if (strncmp(selectedPort, "COM", 3) == 0) {
        int portNum = atoi(selectedPort + 3);
        if (portNum >= 10) {
            sprintf(portName, "\\\\.\\%s", selectedPort);
        } else {
            strcpy(portName, selectedPort);
        }
    } else {
        strcpy(portName, selectedPort);
    }

    // Get baud rate
    HWND hBaud = GetDlgItem(hWndMain, ID_COMBO_BAUDRATE);
    int baudIndex = SendMessageW(hBaud, CB_GETCURSEL, 0, 0);
    wchar_t wbaud[20];
    SendMessageW(hBaud, CB_GETLBTEXT, baudIndex, (LPARAM)wbaud);
    char baudStr[20];
    WideCharToMultiByte(CP_UTF8, 0, wbaud, -1, baudStr, 20, NULL, NULL);
    baudRate = atoi(baudStr);

    // Get data bits
    HWND hDataBits = GetDlgItem(hWndMain, ID_COMBO_DATABITS);
    int dataIndex = SendMessageW(hDataBits, CB_GETCURSEL, 0, 0);
    dataBits = dataIndex + 5;

    // Get parity
    HWND hParity = GetDlgItem(hWndMain, ID_COMBO_PARITY);
    int parityIndex = SendMessageW(hParity, CB_GETCURSEL, 0, 0);
    char parityChar[] = {'N', 'O', 'E', 'M', 'S'};
    parity = parityChar[parityIndex];

    // Get stop bits
    HWND hStopBits = GetDlgItem(hWndMain, ID_COMBO_STOPBITS);
    int stopIndex = SendMessageW(hStopBits, CB_GETCURSEL, 0, 0);
    stopBits = stopIndex == 0 ? 1 : (stopIndex == 1 ? 15 : 2);

    // Get addresses
    char addrStr[20];
    GetWindowTextA(GetDlgItem(hWndMain, ID_EDIT_SLAVE_ADDR), addrStr, 20);
    slaveAddr = atoi(addrStr);

    GetWindowTextA(GetDlgItem(hWndMain, ID_EDIT_MASTER_ADDR), addrStr, 20);
    masterAddr = atoi(addrStr);

    // Create serial port
    port = SerialPort_create(portName, baudRate, dataBits, parity, stopBits);
    if (!port) {
        sprintf(buffer, "Failed to create serial port %s", portName);
        wchar_t wbuffer[256];
        MultiByteToWideChar(CP_UTF8, 0, buffer, -1, wbuffer, 256);
        MessageBoxW(hWndMain, wbuffer, L"Error", MB_ICONERROR);
        return;
    }

    // Create master
    master = CS101_Master_create(port, NULL, NULL, IEC60870_LINK_LAYER_BALANCED);
    if (!master) {
        SerialPort_destroy(port);
        port = NULL;
        MessageBoxW(hWndMain, L"Failed to create master", L"Error", MB_ICONERROR);
        return;
    }

    CS101_Master_setOwnAddress(master, masterAddr);
    CS101_Master_useSlaveAddress(master, slaveAddr);
    CS101_Master_setASDUReceivedHandler(master, asduReceivedHandler, NULL);
    CS101_Master_setLinkLayerStateChanged(master, linkLayerStateChanged, NULL);
    CS101_Master_setRawMessageHandler(master, rawMessageHandler, NULL);

    // Open serial port
    if (!SerialPort_open(port)) {
        CS101_Master_destroy(master);
        SerialPort_destroy(port);
        master = NULL;
        port = NULL;
        sprintf(buffer, "Failed to open serial port %s", portName);
        wchar_t wbuffer[256];
        MultiByteToWideChar(CP_UTF8, 0, buffer, -1, wbuffer, 256);
        MessageBoxW(hWndMain, wbuffer, L"Error", MB_ICONERROR);
        return;
    }

    // Enable/disable controls
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_PORT), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_BAUDRATE), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_DATABITS), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_PARITY), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_STOPBITS), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_EDIT_SLAVE_ADDR), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_EDIT_MASTER_ADDR), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_CONNECT), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_DISCONNECT), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_REFRESH), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_INTERROGATION), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_SEND_CMD), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_TIME_SYNC), TRUE);

    isConnected = true;
    running = true;

    // Start worker thread
    hWorkerThread = CreateThread(NULL, 0, MasterWorkerThread, NULL, 0, &threadId);

    sprintf(buffer, "Connected to %s (%d,%d,%c,%d)", portName, baudRate, dataBits, parity, stopBits);
    wchar_t wbuffer[256];
    MultiByteToWideChar(CP_UTF8, 0, buffer, -1, wbuffer, 256);
    SetWindowTextW(hStatusText, wbuffer);
    AddLog(buffer, FALSE);

    MessageBoxW(hWndMain, L"Connected successfully!", L"Info", MB_ICONINFORMATION);
}

// Disconnect from slave
void OnDisconnect()
{
    if (!isConnected) {
        return;
    }

    running = false;

    // Wait for worker thread
    if (hWorkerThread) {
        WaitForSingleObject(hWorkerThread, 2000);
        CloseHandle(hWorkerThread);
        hWorkerThread = NULL;
    }

    // Destroy master and port
    if (master) {
        CS101_Master_destroy(master);
        master = NULL;
    }

    if (port) {
        SerialPort_close(port);
        SerialPort_destroy(port);
        port = NULL;
    }

    isConnected = false;

    // Enable/disable controls
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_PORT), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_BAUDRATE), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_DATABITS), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_PARITY), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_COMBO_STOPBITS), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_EDIT_SLAVE_ADDR), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_EDIT_MASTER_ADDR), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_CONNECT), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_DISCONNECT), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_REFRESH), TRUE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_INTERROGATION), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_SEND_CMD), FALSE);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_TIME_SYNC), FALSE);

    SetWindowTextW(hStatusText, L"未连接");
    AddLog("Disconnected", FALSE);
}

// Send interrogation command
void OnInterrogation()
{
    if (!isConnected || !master) {
        MessageBoxW(hWndMain, L"Not connected!", L"Error", MB_ICONERROR);
        return;
    }

    char buffer[256];
    sprintf(buffer, "Interrogation: CA=%d QOI=20 (Station)", slaveAddr);
    AddLog(buffer, TRUE);

    CS101_Master_sendInterrogationCommand(master, CS101_COT_ACTIVATION, slaveAddr, IEC60870_QOI_STATION);
}

// Send single command
void OnSendCommand()
{
    if (!isConnected || !master) {
        MessageBoxW(hWndMain, L"Not connected!", L"Error", MB_ICONERROR);
        return;
    }

    char ioaStr[20];
    char valueStr[20];
    GetWindowTextA(GetDlgItem(hWndMain, ID_EDIT_IOA), ioaStr, 20);
    GetWindowTextA(GetDlgItem(hWndMain, ID_EDIT_CMD_VALUE), valueStr, 20);

    int ioa = atoi(ioaStr);
    bool value = (atoi(valueStr) != 0);

    char buffer[256];
    sprintf(buffer, "Single Command: CA=%d IOA=%d Value=%s", slaveAddr, ioa, value ? "ON" : "OFF");
    AddLog(buffer, TRUE);

    InformationObject sc = (InformationObject) SingleCommand_create(NULL, ioa, value, false, 0);
    CS101_Master_sendProcessCommand(master, CS101_COT_ACTIVATION, slaveAddr, sc);
    InformationObject_destroy(sc);
}

// Send time sync command
void OnTimeSync()
{
    if (!isConnected || !master) {
        MessageBoxW(hWndMain, L"Not connected!", L"Error", MB_ICONERROR);
        return;
    }

    struct sCP56Time2a newTime;
    CP56Time2a_createFromMsTimestamp(&newTime, Hal_getTimeInMs());

    char buffer[256];
    sprintf(buffer, "Time Sync: CA=%d", slaveAddr);
    AddLog(buffer, TRUE);

    CS101_Master_sendClockSyncCommand(master, slaveAddr, &newTime);
}

// Clear log
void OnClearLog()
{
    SetWindowTextW(hLogEdit, L"IEC 60870-5-101 Master - Communication Log");
}

// Window procedure
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message) {
        case WM_COMMAND:
            switch (LOWORD(wParam)) {
                case ID_BTN_CONNECT:
                    OnConnect();
                    break;
                case ID_BTN_DISCONNECT:
                    OnDisconnect();
                    break;
                case ID_BTN_REFRESH:
                    OnRefreshPorts();
                    break;
                case ID_BTN_INTERROGATION:
                    OnInterrogation();
                    break;
                case ID_BTN_SEND_CMD:
                    OnSendCommand();
                    break;
                case ID_BTN_TIME_SYNC:
                    OnTimeSync();
                    break;
                case ID_BTN_CLEAR_LOG:
                    OnClearLog();
                    break;
                case IDCANCEL:
                case IDOK:
                    if (isConnected) {
                        OnDisconnect();
                    }
                    PostQuitMessage(0);
                    break;
            }
            break;

        case WM_DESTROY:
            if (isConnected) {
                OnDisconnect();
            }
            PostQuitMessage(0);
            break;

        default:
            return DefWindowProcW(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Main window creation
BOOL CreateMainWindow(HINSTANCE hInstance)
{
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"CS101MasterGUI";
    wc.hIcon = LoadIconW(NULL, IDI_APPLICATION);
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

    if (!RegisterClassW(&wc)) {
        return FALSE;
    }

    hWndMain = CreateWindowExW(
        0,
        L"CS101MasterGUI",
        L"IEC 60870-5-101 Master - GUI",
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
        CW_USEDEFAULT, CW_USEDEFAULT,
        920, 650,
        NULL, NULL, hInstance, NULL
    );

    if (!hWndMain) {
        return FALSE;
    }

    // Serial port settings group
    CreateWindowW(L"STATIC", L"串口设置", WS_CHILD | WS_VISIBLE | SS_LEFT, 10, 10, 200, 20, hWndMain, NULL, hInstance, NULL);

    // Port
    CreateWindowW(L"STATIC", L"串口号:", WS_CHILD | WS_VISIBLE | SS_LEFT, 20, 40, 70, 20, hWndMain, NULL, hInstance, NULL);
    HWND hPort = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, 90, 35, 100, 200, hWndMain, (HMENU)(UINT_PTR)ID_COMBO_PORT, hInstance, NULL);
    InitializePortCombo(hPort);

    // Refresh button
    CreateWindowW(L"BUTTON", L"刷新", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 200, 35, 60, 25, hWndMain, (HMENU)(UINT_PTR)ID_BTN_REFRESH, hInstance, NULL);

    // Baud rate
    CreateWindowW(L"STATIC", L"波特率:", WS_CHILD | WS_VISIBLE | SS_LEFT, 20, 70, 70, 20, hWndMain, NULL, hInstance, NULL);
    HWND hBaud = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, 90, 65, 100, 200, hWndMain, (HMENU)(UINT_PTR)ID_COMBO_BAUDRATE, hInstance, NULL);
    InitializeBaudrateCombo(hBaud);

    // Data bits
    CreateWindowW(L"STATIC", L"数据位:", WS_CHILD | WS_VISIBLE | SS_LEFT, 20, 100, 70, 20, hWndMain, NULL, hInstance, NULL);
    HWND hDataBits = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, 90, 95, 100, 200, hWndMain, (HMENU)(UINT_PTR)ID_COMBO_DATABITS, hInstance, NULL);
    InitializeDataBitsCombo(hDataBits);

    // Parity
    CreateWindowW(L"STATIC", L"校验位:", WS_CHILD | WS_VISIBLE | SS_LEFT, 20, 130, 70, 20, hWndMain, NULL, hInstance, NULL);
    HWND hParity = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, 90, 125, 100, 200, hWndMain, (HMENU)(UINT_PTR)ID_COMBO_PARITY, hInstance, NULL);
    InitializeParityCombo(hParity);

    // Stop bits
    CreateWindowW(L"STATIC", L"停止位:", WS_CHILD | WS_VISIBLE | SS_LEFT, 20, 160, 70, 20, hWndMain, NULL, hInstance, NULL);
    HWND hStopBits = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, 90, 155, 100, 200, hWndMain, (HMENU)(UINT_PTR)ID_COMBO_STOPBITS, hInstance, NULL);
    InitializeStopBitsCombo(hStopBits);

    // Address settings group
    CreateWindowW(L"STATIC", L"地址设置", WS_CHILD | WS_VISIBLE | SS_LEFT, 280, 10, 200, 20, hWndMain, NULL, hInstance, NULL);

    // Slave address
    CreateWindowW(L"STATIC", L"从站地址:", WS_CHILD | WS_VISIBLE | SS_LEFT, 290, 40, 80, 20, hWndMain, NULL, hInstance, NULL);
    CreateWindowW(L"EDIT", L"1", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER, 370, 38, 60, 20, hWndMain, (HMENU)(UINT_PTR)ID_EDIT_SLAVE_ADDR, hInstance, NULL);

    // Master address
    CreateWindowW(L"STATIC", L"主站地址:", WS_CHILD | WS_VISIBLE | SS_LEFT, 290, 70, 80, 20, hWndMain, NULL, hInstance, NULL);
    CreateWindowW(L"EDIT", L"2", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER, 370, 68, 60, 20, hWndMain, (HMENU)(UINT_PTR)ID_EDIT_MASTER_ADDR, hInstance, NULL);

    // Connect buttons
    CreateWindowW(L"BUTTON", L"连接", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 290, 110, 70, 30, hWndMain, (HMENU)(UINT_PTR)ID_BTN_CONNECT, hInstance, NULL);
    CreateWindowW(L"BUTTON", L"断开", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 370, 110, 70, 30, hWndMain, (HMENU)(UINT_PTR)ID_BTN_DISCONNECT, hInstance, NULL);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_DISCONNECT), FALSE);

    // Command settings group
    CreateWindowW(L"STATIC", L"控制命令", WS_CHILD | WS_VISIBLE | SS_LEFT, 480, 10, 200, 20, hWndMain, NULL, hInstance, NULL);

    // IOA
    CreateWindowW(L"STATIC", L"信息对象地址(IOA):", WS_CHILD | WS_VISIBLE | SS_LEFT, 480, 40, 140, 20, hWndMain, NULL, hInstance, NULL);
    CreateWindowW(L"EDIT", L"5000", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER, 620, 38, 80, 20, hWndMain, (HMENU)(UINT_PTR)ID_EDIT_IOA, hInstance, NULL);

    // Command value
    CreateWindowW(L"STATIC", L"命令值(0/1):", WS_CHILD | WS_VISIBLE | SS_LEFT, 480, 70, 100, 20, hWndMain, NULL, hInstance, NULL);
    CreateWindowW(L"EDIT", L"1", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER, 580, 68, 50, 20, hWndMain, (HMENU)(UINT_PTR)ID_EDIT_CMD_VALUE, hInstance, NULL);

    // Command buttons
    CreateWindowW(L"BUTTON", L"遥询命令", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 480, 105, 90, 30, hWndMain, (HMENU)(UINT_PTR)ID_BTN_INTERROGATION, hInstance, NULL);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_INTERROGATION), FALSE);

    CreateWindowW(L"BUTTON", L"发送单点命令", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 580, 105, 120, 30, hWndMain, (HMENU)(UINT_PTR)ID_BTN_SEND_CMD, hInstance, NULL);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_SEND_CMD), FALSE);

    CreateWindowW(L"BUTTON", L"时钟同步", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 480, 145, 90, 30, hWndMain, (HMENU)(UINT_PTR)ID_BTN_TIME_SYNC, hInstance, NULL);
    EnableWindow(GetDlgItem(hWndMain, ID_BTN_TIME_SYNC), FALSE);

    // Status
    hStatusText = CreateWindowW(L"STATIC", L"未连接", WS_CHILD | WS_VISIBLE | SS_LEFT, 10, 195, 880, 20, hWndMain, (HMENU)(UINT_PTR)ID_STATUS_TEXT, hInstance, NULL);

    // Log window
    CreateWindowW(L"STATIC", L"通信日志", WS_CHILD | WS_VISIBLE | SS_LEFT, 10, 225, 200, 20, hWndMain, NULL, hInstance, NULL);

    hLogEdit = CreateWindowW(L"EDIT", L"IEC 60870-5-101 Master - Communication Log", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY, 10, 250, 880, 350, hWndMain, (HMENU)(UINT_PTR)ID_EDIT_LOG, hInstance, NULL);

    // Clear log button
    CreateWindowW(L"BUTTON", L"清空日志", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 800, 605, 90, 30, hWndMain, (HMENU)(UINT_PTR)ID_BTN_CLEAR_LOG, hInstance, NULL);

    return TRUE;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    hInst = hInstance;

    InitializeCriticalSection(&logCS);

    if (!CreateMainWindow(hInstance)) {
        return FALSE;
    }

    ShowWindow(hWndMain, nCmdShow);
    UpdateWindow(hWndMain);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    DeleteCriticalSection(&logCS);

    return (int)msg.wParam;
}
