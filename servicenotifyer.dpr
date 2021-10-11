////////////////////////////////////////////////////////////////////////////////
//
//  ****************************************************************************
//  * Unit Name : SimpleService
//  * Purpose   : ���� ����������� ����������� ������������ �� �������.
//  * Author    : ��������� (Rouse_) ������
//  * Copyright : � Fangorn Wizards Lab 1998 - 2010.
//  * Version   : 1.00
//  ****************************************************************************
//

// ������� �������� ����:
// �.�. ��� ������ �� �������� �������������,
// (����� ����, ������� � Windows Vista, ������� ������� �� ��������������)
// �� �� ����� ���������� ���� ���� �� �������� �������������� ������������,
// ����� ���������� � ���� ������� � ������ ��������.
// ����� �������, ��� � ������� ��� ����� �����������,
// �� ������ ������ ������� � ������ ������� ������ ����������, �����������
// � ��������� �������������� ��������.
// (����� ������� ��� � ����� mailslot, �������� ����� �� \\.\MAILSLOT\MESSNGR,
// �� ������ ������� �� �������������, �.�. �������� ����� ������������ �������
// �� ������ �����������, �� � ����� ��������� ������).
// ����� ������� ���������� ���������� ������
// "������ -> ���������� ������������ ������".
// ������� ���������� ������������ ������ � ���������� �� �����������,
// ��� �� ������� ��������� ������� �������� ������� �� ������
// ������������� ����������, ��� � �� ������� ����, ��� � ������ �������������
// �����������, ������ ���������� ����� ������������� (������� �������������).
// ����� �������, ������ ������ ���������� ��� �� ��� ������� ���������
// ������������ ���������� � ��������� �������� ������������ �
// �������� ��� ������ ��� �����������.
// � ������ � ������������ ���������� ���������� � ����� ����������� �����,
// ������ ������ ���������� ��� ������ ���������� ��������� ������.
// �������� ������� - ��� ����������� ���������� �������� �����
// ������������� ������������ � ������� CreateProcessAsUser
// � ������������ �����������.
// � ������ W2K ��� �������������� ����� ��������� ������
// � ������� ����������� �������, ���� �������� ���������� �� �������� ��������,
// � ��������� ������� (XP � ����) ����� ����� WTSQueryUserToken

program servicenotifyer;

{$DEFINE SERVICE_DEBUG}

uses
  Windows,
  SysUtils,
  WinSvc;

{$WARN SYMBOL_PLATFORM OFF}

const
  InfoStr = 'Use:'#13#10'%s [ -install | -uninstall ]';
  ServiceFileName = 'servicenotifyer.exe';
  Name = 'servicenotifyer';
  DisplayName = 'SimpleService notifyer demo';    

var
  ServicesTable: packed array [0..1] of TServiceTableEntry;
  StatusHandle: SERVICE_STATUS_HANDLE = 0;
  Status: TServiceStatus;

{$R *.res}

// ����� ���������� (��� ������ ������� �� �����������)
// =============================================================================
procedure ShowMsg(Msg: string; Flags: integer = -1);
begin
  if Flags < 0 then Flags := MB_ICONSTOP;
  MessageBox(0, PChar(Msg), ServiceFileName,
    MB_OK or MB_TASKMODAL or MB_TOPMOST or Flags)
end;

// ���������� ������� � SCM
// =============================================================================
function Install: Boolean;
const
  StartType =
{$IFDEF SERVICE_DEBUG}
    SERVICE_DEMAND_START;
{$ELSE}
    SERVICE_AUTO_START;
{$ENDIF}
var
  SCManager, Service: SC_HANDLE;
begin
  SCManager := OpenSCManager(nil, nil, SC_MANAGER_CREATE_SERVICE);
  if SCManager <> 0 then
  try
    Service := CreateService(SCManager, Name, DisplayName, SERVICE_ALL_ACCESS,
      SERVICE_WIN32_OWN_PROCESS, StartType, SERVICE_ERROR_NORMAL,
      PChar('"' + ParamStr(0) + '" -service'), nil, nil, nil, nil, nil);
    if Service <> 0 then
    try
      Result := ChangeServiceConfig(Service, SERVICE_NO_CHANGE,
        SERVICE_NO_CHANGE, SERVICE_NO_CHANGE, nil, nil,
        nil, nil, nil, nil, nil);
    finally
      CloseServiceHandle(Service);
    end
    else
      Result := GetLastError = ERROR_SERVICE_EXISTS;
  finally
    CloseServiceHandle(SCManager);
  end
  else
    Result := False;
end;

// ������������ ������� �� SCM
// =============================================================================
function Uninstall: Boolean;
var
  SCManager, Service: SC_HANDLE;
begin
  SCManager := OpenSCManager(nil, nil, SC_MANAGER_CONNECT);
  if SCManager <> 0 then
  try
    Service := OpenService(SCManager, Name, _DELETE);
    if Service <> 0 then
    try
      Result := DeleteService(Service);
    finally
      CloseServiceHandle(Service);
    end
    else
      Result := GetLastError = ERROR_SERVICE_DOES_NOT_EXIST;
  finally
    CloseServiceHandle(SCManager);
  end
  else
    Result := False;
end;

// ������������� �������
// =============================================================================
function Initialize: Boolean;
begin
  with Status do
  begin
    dwServiceType := SERVICE_WIN32_OWN_PROCESS;
    dwCurrentState := SERVICE_START_PENDING;
    dwControlsAccepted := SERVICE_ACCEPT_STOP or SERVICE_ACCEPT_SHUTDOWN;
    dwWin32ExitCode := NO_ERROR;
    dwServiceSpecificExitCode := 0;
    dwCheckPoint := 1;
    dwWaitHint := 5000
  end;
  Result := SetServiceStatus(StatusHandle, Status);
end;

// ���������� SCM ��� ������ ��������
// =============================================================================
function NotifyIsRunning: Boolean;
begin
  with Status do
  begin
    dwCurrentState := SERVICE_RUNNING;
    dwWin32ExitCode := NO_ERROR;
    dwCheckPoint := 0;
    dwWaitHint := 0
  end;
  Result := SetServiceStatus(StatusHandle, Status);
end;

// ���������� ������ �������
// =============================================================================
procedure Stop(Code: DWORD = NO_ERROR);
begin
  with Status do
  begin
    dwCurrentState := SERVICE_STOPPED;
    dwWin32ExitCode := Code;
  end;
  SetServiceStatus(StatusHandle, Status);
end;    

// ����� ��� ������� � ����� �������� �������� SCM
// =============================================================================
function ServicesCtrlHandler(dwControl: DWORD): DWORD; stdcall;
begin
  Result := 1;
  case dwControl of
    SERVICE_CONTROL_STOP, SERVICE_CONTROL_SHUTDOWN:
      Stop;
    SERVICE_CONTROL_INTERROGATE:
      NotifyIsRunning;
  end;
end;

type
  TWTSQueryUserToken = function(
    SessionID: DWORD; var Token: THandle): BOOL; stdcall;

  function WTSGetActiveConsoleSessionId: DWORD; stdcall;
    external 'kernel32.dll';

//  ������� ���������� � ������ ������� ������� ��� W2K.
//  ��� ������ �������� ����� ������ ���������� ��������,
//  ���� �������� ������� � ������ ��������� ��� ������������ ��������.
// =============================================================================
function EnumDesktopWindowsCallback(
  WndHandle: THandle; Param: LPARAM): BOOL; stdcall;
var
  ProcessID: DWORD;
  ProcessHandle, UserToken: THandle;
begin
  Result := True;
  GetWindowThreadProcessId(WndHandle, ProcessID);
  ProcessHandle := OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID);
  if ProcessHandle <> 0 then
  try
    if OpenProcessToken(ProcessHandle, TOKEN_ALL_ACCESS, UserToken) then
    begin
      PDWORD(Param)^ := UserToken;
      Result := False;
    end;
  finally
    CloseHandle(ProcessHandle);
  end;
end;

//  ��������������� ������ ������������� ����������,
//  � ��������� �������������� ��������
// =============================================================================
function ShowNotify(const Value: string): DWORD;
const
  WINDOW_STATION_NAME = 'Winsta0';
  APPLICATION_DESKTOP_NAME = 'Default';
var
  hLib: THandle;
  hCurrentWinStation, hInteractiveWorkstation: HWINSTA;
  hDefaultDesktop: HDESK;
  SI: TStartupInfo;
  PI: TProcessInformation;
  SessionId: DWORD;
  hInteractiveToken: THandle;
  WTSQueryUserToken: TWTSQueryUserToken;
begin
  Result := NO_ERROR;
  hInteractiveToken := INVALID_HANDLE_VALUE;
  if (Win32MajorVersion = 5) and (Win32MinorVersion = 0) then
  begin
    // � ������ W2K
    hCurrentWinStation := GetProcessWindowStation;
    // ��������� ������� ������� ������������
    hInteractiveWorkstation := OpenWindowStation(
      PChar(WINDOW_STATION_NAME), False, MAXIMUM_ALLOWED);
    if hInteractiveWorkstation = 0 then Exit;
    try
      // ���������� � ��� ��� �������
      if not SetProcessWindowStation(hInteractiveWorkstation) then Exit;
      try
        // ��������� ������������� �������
        hDefaultDesktop := OpenDesktop(PChar(APPLICATION_DESKTOP_NAME),
          0, False, MAXIMUM_ALLOWED);
        if hDefaultDesktop = 0 then Exit;
        try
          // ����������� ���� �������� � ����� �������
          // ����� ������������� ������������
          EnumDesktopWindows(hDefaultDesktop, @EnumDesktopWindowsCallback,
            Integer(@hInteractiveToken));
        finally
          CloseDesktop(hDefaultDesktop);
        end;
      finally
        SetProcessWindowStation(hCurrentWinStation);
      end;
    finally
      CloseWindowStation(hInteractiveWorkstation);
    end;
  end
  else
  begin
    // � ������ Windows �� � ���� ���������� ����������
    hLib := LoadLibrary('Wtsapi32.dll');
    if hLib > HINSTANCE_ERROR then
    begin
      // �������� ����� ������� WTSQueryUserToken
      @WTSQueryUserToken := GetProcAddress(hLib, 'WTSQueryUserToken');
      if Assigned(@WTSQueryUserToken) then
      begin
        // �������� ID ������ � ������ �������
        // ����� ������ ����������� ������������
        SessionID := WTSGetActiveConsoleSessionId;
        // �������� ����� ������������
        WTSQueryUserToken(SessionID, hInteractiveToken);
      end;
    end;
  end;
  if hInteractiveToken = INVALID_HANDLE_VALUE then
  begin
    Result := GetLastError;
    Exit;
  end;
  // ����� ���� ��� ����� ������� - ���������� ������ ������ ����
  // � ���������� notify � �����������, ������� ���������� ����������
  try
    ZeroMemory(@SI, SizeOf(TStartupInfo));
    SI.cb := SizeOf(TStartupInfo);
    SI.lpDesktop := PChar(WINDOW_STATION_NAME + '\' +
      APPLICATION_DESKTOP_NAME);
    if not CreateProcessAsUser(hInteractiveToken,
      PChar(ParamStr(0)),
      PChar('"' + ParamStr(0) + '" -notify ' + Value), nil, nil, False,
      NORMAL_PRIORITY_CLASS, nil, nil, SI, PI) then
      Result := GetLastError;
  finally
    CloseHandle(hInteractiveToken);
  end;
end;

// ������� ��������� �������
// =============================================================================
procedure MainProc(ArgCount: DWORD; var Args: array of PChar); stdcall;
var
  I: Integer;
  dwResult, dwDelay: DWORD;
begin
  StatusHandle := RegisterServiceCtrlHandler(Name, @ServicesCtrlHandler);
  if (StatusHandle <> 0) and Initialize and NotifyIsRunning then
  begin
    dwResult := NO_ERROR;
    while Status.dwCurrentState <> SERVICE_STOP do
    try
      try
        Randomize;
        for I := 1 to 3 do
        begin
          dwDelay := Random(10) + 1;
          dwResult := ShowNotify(
            '����������� �� ������� ��������� ������������ �' +
            IntToStr(I) + ' \n ' +
            '��������� ����������� ����� ' + IntToStr(dwDelay) + ' ���.');
          Sleep(dwDelay * 1000);
          if dwResult <> NO_ERROR then
            Break;
        end;
        dwResult := ShowNotify(
          '���� ����������� ��������, ������ ��������� ���� ������.');
      finally
        Stop(dwResult);
      end;
    except
      // ��������� ������ �������
      on E: EOSError do
      begin
        Stop(E.ErrorCode);
      end;
    end;
  end;
end;

var
  NotifyString: string;
  I: Integer;
begin
  if ParamCount > 0 then
  begin
    // ����������
    if AnsiUpperCase(ParamStr(1)) = '-INSTALL' then
    begin
      if not Install then ShowMsg(SysErrorMessage(GetLastError));
      Exit;
    end;
    // ������������
    if AnsiUpperCase(ParamStr(1)) = '-UNINSTALL' then
    begin
      if not Uninstall then ShowMsg(SysErrorMessage(GetLastError));
      Exit;
    end;
    // ������ �������
    if AnsiUpperCase(ParamStr(1)) = '-SERVICE' then
    begin
      ServicesTable[0].lpServiceName := Name;
      ServicesTable[0].lpServiceProc := @MainProc;
      ServicesTable[1].lpServiceName := nil;
      ServicesTable[1].lpServiceProc := nil;
      // ��������� ������, ������ ������ ���� � ������� ���������
      if not StartServiceCtrlDispatcher(ServicesTable[0]) and
        (GetLastError <> ERROR_SERVICE_ALREADY_RUNNING) then
          ShowMsg(SysErrorMessage(GetLastError));
      Exit;
    end;
    // ������ � ������ ������������� ����������
    if AnsiUpperCase(ParamStr(1)) = '-NOTIFY' then
    begin
      NotifyString := '';
      for I := 2 to ParamCount do
        if ParamStr(I) = '\n' then
          NotifyString := NotifyString + sLineBreak
        else
          NotifyString := NotifyString + ParamStr(I) + ' ';
      MessageBox(0, PChar(Trim(NotifyString)), '����������� �� �������',
        MB_ICONINFORMATION);
    end
    else
      ShowMsg(Format(InfoStr, [ServiceFileName]), MB_ICONINFORMATION);
  end
  else
    ShowMsg(Format(InfoStr, [ServiceFileName]), MB_ICONINFORMATION);
end.
