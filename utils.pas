unit utils;

interface
uses
  Winapi.Windows, fsplugin, Winapi.ShellApi;

type
  TStringArray = array of string;

  TSearchRec = record
  private
    function GetTimeStamp: TDateTime;
  public
{$IFDEF MSWINDOWS}
    Time: Integer platform deprecated;
{$ENDIF MSWINDOWS}
{$IFDEF POSIX}
    Time: time_t platform;
{$ENDIF POSIX}
    Size: Int64;
    Attr: Integer;
    Name: string;
    ExcludeAttr: Integer;
{$IFDEF MSWINDOWS}
    FindHandle: THandle platform;
    FindData: TWin32FindData platform;
{$ENDIF MSWINDOWS}
{$IFDEF POSIX}
    Mode: mode_t platform;
    FindHandle: Pointer platform;
    PathOnly: string platform;
    Pattern: string platform;
{$ENDIF POSIX}
    property TimeStamp: TDateTime read GetTimeStamp;
  end;

  WordRec = packed record
    case Integer of
      0: (Lo, Hi: Byte);
      1: (Bytes: array [0..1] of Byte);
  end;

  LongRec = packed record
    case Integer of
      0: (Lo, Hi: Word);
      1: (Words: array [0..1] of Word);
      2: (Bytes: array [0..3] of Byte);
  end;

  Int64Rec = packed record
    case Integer of
      0: (Lo, Hi: Cardinal);
      1: (Cardinals: array [0..1] of Cardinal);
      2: (Words: array [0..3] of Word);
      3: (Bytes: array [0..7] of Byte);
  end;



function FileExists(const FileName: string; FollowLink: Boolean = True): Boolean;
function ExtractPathItem(var S: string): string;
function Min(X, Y: NativeInt): NativeInt; {$IFNDEF DEBUG}inline;{$ENDIF}
function Trim(const S: string): string;
function StrReplace(const S: string; AFrom, ATo: Char): string;
function IIF(ACondition: Boolean; ATrue, AFalse: NativeUInt): NativeUInt; inline; overload;
function IIF(ACondition: Boolean; const ATrue, AFalse: string): string; inline; overload;
function ChangeFileExt(const FileName, Extension: string): string;
function LastDelimiter(const Delimiters, S: string): Integer;
function StrScan(const Str: PWideChar; Chr: WideChar): PWideChar; overload;
function ExtractFilePath(const FileName: string): string;
function ExtractFileName(const FileName: string): string;
function GetModuleName(Module: HMODULE): string;
function SplitStr(
  const AString: string; var AName, AValue: string;
  const ADelimiter: string = '='): Boolean;
function GetFileList(
  const Dir: string; const Mask: string = '*.*';
  IncludeSubDirs: Boolean = False): TStringArray;
procedure DynArrayAppend(var A: Pointer; const Append: Pointer; ItemsTypeInfo: Pointer); overload;
procedure DynArrayAppend(var A: TStringArray; const Append: TStringArray); overload;
function LoadFileData(const AFileName: string; var AValue: string): NativeInt;
function SetVarAsAdmin(
  const AVarName, ASourceFile, AValue, ANewName: string): Cardinal;
function SaveToFile(const AFileName, AValue: string): NativeInt;
function IntToStr(Value: Integer): string;
function CreateDefaultSecurityAttributes(Inheritable: Boolean): TSecurityAttributes;


const
{ File attribute constants }

  faInvalid     = -1;
  faReadOnly    = $00000001;
  faHidden      = $00000002 platform; // only a convention on POSIX
  faSysFile     = $00000004 platform; // on POSIX system files are not regular files and not directories
  faVolumeID    = $00000008 platform deprecated;  // not used in Win32
  faDirectory   = $00000010;
  faArchive     = $00000020 platform;
  faNormal      = $00000080;
  faTemporary   = $00000100 platform;
  faSymLink     = $00000400 platform; // Available on POSIX and Vista and above
  faCompressed  = $00000800 platform;
  faEncrypted   = $00004000 platform;
  faVirtual     = $00010000 platform;
  faAnyFile     = $000001FF;

  HoursPerDay   = 24;
  MinsPerHour   = 60;
  SecsPerMin    = 60;
  MSecsPerSec   = 1000;
  MinsPerDay    = HoursPerDay * MinsPerHour;
  SecsPerDay    = MinsPerDay * SecsPerMin;
  SecsPerHour   = SecsPerMin * MinsPerHour;
  MSecsPerDay   = SecsPerDay * MSecsPerSec;


type
  TTimeStamp = record
    Time: Integer;      { Number of milliseconds since midnight }
    Date: Integer;      { One plus number of days since 1/1/0001 }
  end;

  PDayTable = ^TDayTable;
  TDayTable = array[1..12] of Word;

{ The MonthDays array can be used to quickly find the number of
  days in a month:  MonthDays[IsLeapYear(Y), M]      }

const
  MonthDays: array [Boolean] of TDayTable =
    ((31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31),
     (31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31));

  DateDelta = 693594;



type
  TCommandRec = packed record
    Typ: Byte;
    ByteCount: LONG;
  end;

  TCommand = (tUnknown, tSetValue, tRename);


implementation



function FileExists(const FileName: string; FollowLink: Boolean = True): Boolean;

  function ExistsLockedOrShared(const Filename: string): Boolean;
  var
    FindData: TWin32FindData;
    LHandle: THandle;
  begin
    { Either the file is locked/share_exclusive or we got an access denied }
    LHandle := FindFirstFile(PChar(Filename), FindData);
    if LHandle <> INVALID_HANDLE_VALUE then
    begin
      Winapi.Windows.FindClose(LHandle);
      Result := FindData.dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY = 0;
    end
    else
      Result := False;
  end;

var
  Flags: Cardinal;
  Handle: THandle;
  LastError: Cardinal;
begin
  Flags := GetFileAttributes(PChar(FileName));

  if Flags <> INVALID_FILE_ATTRIBUTES then
  begin
{$WARNINGS OFF}
    if faSymLink and Flags <> 0 then
{$WARNINGS ON}
    begin
      if not FollowLink then
        Exit(True)
      else
      begin
        if faDirectory and Flags <> 0 then
          Exit(False)
        else
        begin
          Handle := CreateFile(PChar(FileName), GENERIC_READ, FILE_SHARE_READ, nil,
            OPEN_EXISTING, 0, 0);
          if Handle <> INVALID_HANDLE_VALUE then
          begin
            CloseHandle(Handle);
            Exit(True);
          end;
          LastError := GetLastError;
          Exit(LastError = ERROR_SHARING_VIOLATION);
        end;
      end;
    end;

    Exit(faDirectory and Flags = 0);
  end;

  LastError := GetLastError;
  Result := (LastError <> ERROR_FILE_NOT_FOUND) and
    (LastError <> ERROR_PATH_NOT_FOUND) and
    (LastError <> ERROR_INVALID_NAME) and ExistsLockedOrShared(Filename);
end;

function ExtractPathItem(var S: string): string;
var
  i, start, l: Integer;
begin
  l := Length(S);
  if (l > 0) then
  begin
    if (S[1] = '\') then
      start := 2
    else
      start := 1;
    for i := start to l do
    begin
      if (S[i] = '\') then
      begin
        Result := Copy(S, start, i - start);
        S := Copy(S, i + 1, l);
        Break;
      end
      else if (i = l) then
      begin
        Result := Copy(S, start, i - start + 1);
        S := '';
      end;
    end;
  end
  else
    Result := '';
end;

function Min(X, Y: NativeInt): NativeInt; {$IFNDEF DEBUG}inline;{$ENDIF}
begin
  if (Y < X) then
    Result := Y
  else
    Result := X;
end;

function Trim(const S: string): string;
var
  I, L: Integer;
begin
  L := Length(S);
  I := 1;
  if (L > 0) and (S[I] > ' ') and (S[I] <> #$FEFF) and (S[L] > ' ') then Exit(S);
  while (I <= L) and ((S[I] <= ' ') or (S[I] = #$FEFF)) do Inc(I);
  if I > L then Exit('');
  while S[L] <= ' ' do Dec(L);
  Result := Copy(S, I, L - I + 1);
end;

function StrReplace(const S: string; AFrom, ATo: Char): string;
var
  i: Integer;
begin
  Result := S;
  for i := 1 to Length(Result) do
    if (Result[i] = AFrom) then
      Result[i] := ATo;
end;

function IIF(ACondition: Boolean; ATrue, AFalse: NativeUInt): NativeUInt; inline;
begin
  if (ACondition) then
    Result := ATrue
  else
    Result := AFalse;
end;

function IIF(ACondition: Boolean; const ATrue, AFalse: string): string; inline;
begin
  if (ACondition) then
    Result := ATrue
  else
    Result := AFalse;
end;

function ChangeFileExt(const FileName, Extension: string): string;
var
  I: Integer;
begin
  I := LastDelimiter('.\:', Filename);
  if (I = 0) or (FileName[I] <> '.') then I := MaxInt;
  Result := Copy(FileName, 1, I - 1) + Extension;
end;

function LastDelimiter(const Delimiters, S: string): Integer;
var
  P: PChar;
begin
  Result := Length(S);
  P := PChar(Delimiters);
  while Result > 0 do
  begin
    if (S[Result] <> #0) and (StrScan(P, S[Result]) <> nil) then
      Exit;
    Dec(Result);
  end;
end;

function StrScan(const Str: PWideChar; Chr: WideChar): PWideChar;
begin
  Result := Str;
  while Result^ <> #0 do
  begin
    if Result^ = Chr then
      Exit;
    Inc(Result);
  end;
  if Chr <> #0 then
    Result := nil;
end;

function ExtractFilePath(const FileName: string): string;
var
  I: Integer;
begin
  I := LastDelimiter('\:', FileName);
  Result := Copy(FileName, 1, I);
end;

function ExtractFileName(const FileName: string): string;
var
  I: Integer;
begin
  I := LastDelimiter('\:', FileName);
  Result := Copy(FileName, I + 1, MaxInt);
end;

function GetModuleName(Module: HMODULE): string;
var
  ModName: array[0..MAX_PATH] of Char;
begin
  SetString(Result, ModName, GetModuleFileName(Module, ModName, Length(ModName)));
end;

function SplitStr(
  const AString: string; var AName, AValue: string;
  const ADelimiter: string): Boolean;
var
  delPos: Integer;
begin
  delPos := Pos(ADelimiter, AString);
  Result := (delPos > 0);
  if (Result) then
  begin
    AName := Copy(AString, 1, delPos - 1);
    AValue := Copy(AString, delPos + 1, Length(AString));
  end;
end;

{$WARNINGS OFF}
function FindMatchingFile(var F: TSearchRec): Integer;
var
  LocalFileTime: TFileTime;
begin
  with F do
  begin
    while FindData.dwFileAttributes and ExcludeAttr <> 0 do
      if not FindNextFile(FindHandle, FindData) then
      begin
        Result := GetLastError;
        Exit;
      end;
    FileTimeToLocalFileTime(FindData.ftLastWriteTime, LocalFileTime);
    FileTimeToDosDateTime(LocalFileTime, LongRec(Time).Hi,
      LongRec(Time).Lo);
    Size := FindData.nFileSizeLow or Int64(FindData.nFileSizeHigh) shl 32;
    Attr := FindData.dwFileAttributes;
    Name := FindData.cFileName;
  end;
  Result := 0;
end;

procedure FindClose(var F: TSearchRec);
begin
{$IFDEF MSWINDOWS}
  if F.FindHandle <> INVALID_HANDLE_VALUE then
  begin
    Winapi.Windows.FindClose(F.FindHandle);
    F.FindHandle := INVALID_HANDLE_VALUE;
  end;
{$ENDIF MSWINDOWS}
{$IFDEF POSIX}
  if F.FindHandle <> nil then
  begin
    closedir(F.FindHandle);
    F.FindHandle := nil;
  end;
{$ENDIF POSIX}
end;

function FindFirst(const Path: string; Attr: Integer;
  var F: TSearchRec): Integer;
const
  faSpecial = faHidden or faSysFile or faDirectory;
begin
  F.ExcludeAttr := not Attr and faSpecial;
  F.FindHandle := FindFirstFile(PChar(Path), F.FindData);
  if F.FindHandle <> INVALID_HANDLE_VALUE then
  begin
    Result := FindMatchingFile(F);
    if Result <> 0 then FindClose(F);
  end
  else
    Result := GetLastError;
end;

function FindNext(var F: TSearchRec): Integer;
begin
{$IFDEF MSWINDOWS}
  if FindNextFile(F.FindHandle, F.FindData) then
    Result := FindMatchingFile(F)
  else
    Result := GetLastError;
{$ENDIF MSWINDOWS}
{$IFDEF POSIX}
  Result := FindMatchingFile(F);
{$ENDIF POSIX}
end;
{$WARNINGS ON}

procedure FreeUniArray(var A: TStringArray);
{$IF RTLVersion < 20.00}
var
  i: Integer;
{$IFEND}
begin
  {$IF RTLVersion < 20.00}
  for i := Length(A) - 1 downto 0 do
    A[i] := '';
  {$IFEND}
  SetLength(A, 0);
end;

procedure DynArrayAddItem(var A: TStringArray; const Item: string);
var
  l: Integer;
begin
  l := Length(TStringArray(A));
  SetLength(TStringArray(A), l + 1);
  TStringArray(A)[l] := Item;
end;

procedure DynArrayAppend(var A: Pointer; const Append: Pointer; ItemsTypeInfo: Pointer); overload;
var
  l, la: Integer;
begin
  la := DynArraySize(Append);
  if (la > 0) then
  begin
    l := DynArraySize(A);
    SetLength(TBoundArray(A), l + la);
    if (ItemsTypeInfo <> nil) then
      CopyArray(@TBoundArray(A)[l], @TBoundArray(Append)[0], ItemsTypeInfo, la)
    else
      Move(TBoundArray(Append)[0], TBoundArray(A)[l], la * SizeOf(TBoundArray(A)[0]));
  end;
end;

procedure DynArrayAppend(var A: TStringArray; const Append: TStringArray); overload;
begin
  DynArrayAppend(Pointer(A), Pointer(Append), TypeInfo(TStringArray));
end;

function GetFileList(
  const Dir: string; const Mask: string; IncludeSubDirs: Boolean): TStringArray;
var
  sr: TSearchRec;
  r, i: Integer;
  s, fn: string;
  subDirItems: TStringArray;
begin
  SetLength(Result, 0);
  if (Dir <> '') then
  begin
    s := Dir;
    if (s[Length(s)] <> '\') then
      s := s + '\';
    r := FindFirst(s + Mask, faAnyFile and not faDirectory, sr);
    while (r = 0) do
    begin
      if (sr.Attr and faDirectory <> faDirectory) then
      begin
        fn := s + sr.Name;
        DynArrayAddItem(Result, sr.Name);
      end;
      r := FindNext(sr);
    end;
    FindClose(sr);

    if IncludeSubDirs then
    begin
      r := FindFirst(s + '*.*', faDirectory, sr);
      while (r = 0) do
      begin
        if (sr.Attr and faDirectory = faDirectory) and (sr.Name <> '.') and (sr.Name <> '..') then
        begin
          fn := s + sr.Name;
          subDirItems := GetFileList(fn, Mask, IncludeSubDirs);
          if (Length(subDirItems) > 0) then
          begin
            for i := 0 to Length(subDirItems) - 1 do
              subDirItems[i] := sr.Name + '\' + subDirItems[i];
            DynArrayAppend(Result, subDirItems);
          end;
        end;
        r := FindNext(sr);
      end;
      FindClose(sr);
    end;
  end;
end;

function IsLeapYear(Year: Word): Boolean;
begin
  Result := (Year mod 4 = 0) and ((Year mod 100 <> 0) or (Year mod 400 = 0));
end;

function TryEncodeDate(Year, Month, Day: Word; out Date: TDateTime): Boolean;
var
  I: Integer;
  DayTable: PDayTable;
begin
  Result := False;
  DayTable := @MonthDays[IsLeapYear(Year)];
  if (Year >= 1) and (Year <= 9999) and (Month >= 1) and (Month <= 12) and
    (Day >= 1) and (Day <= DayTable^[Month]) then
  begin
    for I := 1 to Month - 1 do Inc(Day, DayTable^[I]);
    I := Year - 1;
    Date := I * 365 + I div 4 - I div 100 + I div 400 + Day - DateDelta;
    Result := True;
  end;
end;

function TimeStampToDateTime(const TimeStamp: TTimeStamp): TDateTime;
var
  Temp: Int64;
  ms: Single;
begin
//  ValidateTimeStamp(TimeStamp);
  Temp := TimeStamp.Date;
  Dec(Temp, DateDelta);
  Temp := Temp * Integer(MSecsPerDay);

  if Temp >= 0 then
    Inc(Temp, TimeStamp.Time)
  else
    Dec(Temp, TimeStamp.Time);

  ms := MSecsPerDay;
  Result := Temp / ms;
end;

function TryEncodeTime(Hour, Min, Sec, MSec: Word; out Time: TDateTime): Boolean;
var
  TS: TTimeStamp;
begin
  Result := False;
  if (Hour < HoursPerDay) and (Min < MinsPerHour) and (Sec < SecsPerMin) and (MSec < MSecsPerSec) then
  begin
    TS.Time :=  (Hour * (MinsPerHour * SecsPerMin * MSecsPerSec))
              + (Min * SecsPerMin * MSecsPerSec)
              + (Sec * MSecsPerSec)
              +  MSec;
    TS.Date := DateDelta; // This is the "zero" day for a TTimeStamp, days between 1/1/0001 and 12/30/1899 including the latter date
    Time := TimeStampToDateTime(TS);
    Result := True;
  end;
end;

function InternalFileTimeToDateTime(Time: TFileTime): TDateTime;

  function InternalEncodeDateTime(const AYear, AMonth, ADay, AHour, AMinute, ASecond,
    AMilliSecond: Word): TDateTime;
  var
    LTime: TDateTime;
    Success: Boolean;
  begin
    Result := 0;
    Success := TryEncodeDate(AYear, AMonth, ADay, Result);
    if Success then
    begin
      Success := TryEncodeTime(AHour, AMinute, ASecond, AMilliSecond, LTime);
      if Success then
        if Result >= 0 then
          Result := Result + LTime
        else
          Result := Result - LTime
    end;
  end;

var
  LFileTime: TFileTime;
  SysTime: TSystemTime;
begin
  Result := 0;
  FileTimeToLocalFileTime(Time, LFileTime);

  if FileTimeToSystemTime(LFileTime, SysTime) then
    with SysTime do
    begin
      Result := InternalEncodeDateTime(wYear, wMonth, wDay, wHour, wMinute,
        wSecond, wMilliseconds);
    end;
end;


{ TSearchRec }

function TSearchRec.GetTimeStamp: TDateTime;
begin
{$IFDEF MSWINDOWS}
{$WARNINGS OFF}
  Result := InternalFileTimeToDateTime(FindData.ftLastWriteTime);
{$WARNINGS ON}
{$ENDIF}
{$IFDEF POSIX}
  Result := InternalFileTimeToDateTime(Time);
{$ENDIF POSIX}
end;

function LoadFileData(const AFileName: string; var AValue: string): NativeInt;
var
  line: string;
  line2: AnsiString;
  f: THandle;
  bw: Cardinal;
  l, delPos: Integer;
  bom: WideChar;
begin
  f :=
    CreateFile(
      PChar(AFileName), GENERIC_READ, FILE_SHARE_READ, nil, OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL, 0);
  if (f <> INVALID_HANDLE_VALUE) then
  begin
    try
      l := Min(4096 * 2, GetFileSize(f, nil) div 2 * 2);
      bom := #0;
      ReadFile(f, bom, SizeOf(bom), bw, nil);
      if (bom = #$FEFF) then
      begin
        SetLength(line, l div 2);
        bw := 0;
        ReadFile(f, line[1], l, bw, nil);
        SetLength(line, bw div 2);
      end
      else
      begin
        SetFilePointer(f, 0, nil, FILE_BEGIN);
        SetLength(line2, l);
        bw := 0;
        ReadFile(f, line2[1], l, bw, nil);
        SetLength(line, bw);
        line := string(line2);
      end;
    finally
      CloseHandle(f);
    end;

    delPos := Pos(#13, line);
    if (delPos > 0) then
      line := Copy(line, 1, delPos - 1);
    AValue := Trim(line);
    Result := ERROR_SUCCESS;
  end
  else
    Result := ERROR_ACCESS_DENIED;
end;

function EnumWindowsProc(hwnd: HWND; lParam: PInteger): BOOL; stdcall;
var
  pid: Cardinal;
begin
  Result := True;
  pid := GetWindowThreadProcessId(hwnd, nil);
  if (pid = GetCurrentProcess) then
  begin
    lParam^ := hwnd;
    Result := False;
  end;
end;


function SetVarAsAdmin(
  const AVarName, ASourceFile, AValue, ANewName: string): Cardinal;

  function getName: string;
  begin
    Result := IntToStr(Random(MaxInt));
  end;

  function getTempFile: string;
  var
    s: string;
  begin
    SetLength(Result, MAX_PATH);
    SetLength(Result, GetTempPath(Length(Result), PChar(Result)));
    if (Result <> '') then
    begin
      if (Result[Length(Result)] <> '\') then
        Result := Result + '\';
      repeat
        s := Result + getName;
      until not FileExists(s);
      Result := s;
      if (SaveToFile(Result, AValue) <> ERROR_SUCCESS) then
        Result := '';
    end
    else
      Result := '';
  end;

var
  sei: TShellExecuteInfo;
  wnd: Integer;
  elevator, source: string;
  vi: TOSVersionInfo;
  pipe: THandle;
  connected: Boolean;
  code: LONG;
  br: Cardinal;
  pipeSec: TSecurityAttributes;
  cmdType: TCommand;

  procedure writeParamCount(ACnt: LONG);
  var
    cmd: TCommandRec;
  begin
    cmd.Typ := 1;
    cmd.ByteCount := ACnt;
    WriteFile(pipe, cmd, SizeOf(cmd), br, nil);
  end;

  procedure writeParam(AParam: TCommand); overload;
  var
    cmd: TCommandRec;
  begin
    cmd.Typ := 2;
    cmd.ByteCount := Ord(AParam);
    WriteFile(pipe, cmd, SizeOf(cmd), br, nil);
  end;

  procedure writeParam(const AParam: string); overload;
  var
    cmd: TCommandRec;
  begin
    cmd.Typ := 3;
    cmd.ByteCount := Length(AParam) * SizeOf(Char);
    WriteFile(pipe, cmd, SizeOf(cmd), br, nil);
    WriteFile(pipe, PChar(AParam)^, cmd.ByteCount, br, nil);
  end;

begin
  FillChar(vi, SizeOf(vi), 0);
  vi.dwOSVersionInfoSize := SizeOf(vi);
  GetVersionEx(vi);
  if (vi.dwMajorVersion >= 6) then
  begin
    elevator := ExtractFilePath(GetModuleName(HInstance)) + 'elevate.exe';
    if (FileExists(elevator)) then
    begin
      if (ANewName <> '') then
      begin
        // rename
        cmdType := tRename;
      end
      else
      begin
        // set value
        source := ASourceFile;
        if (source = '') then
          source := getTempFile;
        cmdType := tSetValue;
      end;

      if (source <> '') or (ANewName <> '') then
      begin
        wnd := 0;
        EnumWindows(@EnumWindowsProc, Integer(@wnd));
        pipeSec := CreateDefaultSecurityAttributes(False);
        pipe :=
          CreateNamedPipe(
            '\\.\pipe\EnvVarsElevateGate', PIPE_ACCESS_DUPLEX,
            PIPE_TYPE_BYTE or PIPE_READMODE_BYTE or PIPE_WAIT,
            PIPE_UNLIMITED_INSTANCES,
            4, 4, 0, nil);
        if (pipe <> INVALID_HANDLE_VALUE) then
        begin
          FillChar(sei, Sizeof(sei), 0);
          sei.cbSize := SizeOf(sei);
          sei.fMask := SEE_MASK_FLAG_DDEWAIT or SEE_MASK_FLAG_NO_UI;
          sei.Wnd := wnd;
          sei.lpVerb := 'runas';
          sei.lpFile := PChar(elevator);
          sei.lpParameters := PChar(IntToStr(GetCurrentProcessId));
          sei.nShow := SW_SHOWNORMAL;

          if (ShellExecuteEx(@sei)) then
          begin
            connected := ConnectNamedPipe(pipe, nil);
//            if (not connected) then
//              connected := (GetLastError = ERROR_NO_DATA); // client already disconnected, but we can read
            if (connected) then
            begin
              writeParamCount(3);
              writeParam(cmdType);
              if (cmdType = tSetValue) then
              begin
                writeParam(AVarName);
                writeParam(source);
              end
              else
              begin
                writeParam(AVarName);
                writeParam(ANewName);
              end;

              if (ReadFile(pipe, code, SizeOf(code), br, nil) and (br = SizeOf(code))) then
              begin
                Result := code;
              end
              else
              begin
                Result := GetLastError;
                if (Result = ERROR_SUCCESS) then
                  Result := ERROR_INVALID_FUNCTION;
              end;
            end
            else
              Result := ERROR_PIPE_NOT_CONNECTED;
          end
          else
            Result := GetLastError;
          if (ASourceFile = '') and (source <> '') then
            DeleteFile(PChar(source));
          CloseHandle(pipe);
        end
        else
          Result := GetLastError;
      end
      else
        Result := ERROR_PATH_NOT_FOUND;
    end
    else
      Result := ERROR_FILE_NOT_FOUND;
  end
  else
    Result := ERROR_ACCESS_DENIED;
end;

function SaveToFile(const AFileName, AValue: string): NativeInt;
var
  f: THandle;
  data: string;
  bw: Cardinal;
begin
  f :=
    CreateFile(
      PChar(AFileName), GENERIC_READ or GENERIC_WRITE, FILE_SHARE_READ, nil,
      CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
  if (f <> INVALID_HANDLE_VALUE) then
  begin
    try
      data := #$FEFF + AValue;
      WriteFile(f, data[1], Length(data) * SizeOf(Char), bw, nil);
      Result := ERROR_SUCCESS;
    finally
      CloseHandle(f);
    end;
  end
  else
    Result := GetLastError;
end;

function IntToStr( Value : Integer ) : string;
var Buf : Array[ 0..15 ] of Char;
    Dst : PChar;
    Minus : Boolean;
    D: DWORD;
begin
  Dst := @Buf[ 15 ];
  Dst^ := #0;
  Minus := False;
  if Value < 0 then
  begin
    Value := -Value;
    Minus := True;
  end;
  D := Value;
  repeat
    Dec( Dst );
    Dst^ := Char( (D mod 10) + Byte( '0' ) );
    D := D div 10;
  until D = 0;
  if Minus then
  begin
    Dec( Dst );
    Dst^ := '-';
  end;
  Result := Dst;
end;

function CreateDefaultSecurityAttributes(Inheritable: Boolean): TSecurityAttributes;
var
  pSD: PSECURITY_DESCRIPTOR;
begin
  Result.nLength := 0;
  // —оздать пустой дескриптор безопасности, позвол€ющий всем писать в канал.
  // ѕредупреждение: ”казание nil в качестве последнего параметра функции
  // CreateNamedPipe() означает, что все клиенты, подсоединившиес€ к каналу
  // будут иметь те же атрибуты безопасности, что и пользователь, чь€ учетна€
  // запись использовалась при создании серверной стороны канала.
  pSD := PSECURITY_DESCRIPTOR(LocalAlloc(LPTR, SECURITY_DESCRIPTOR_MIN_LENGTH));
  if not Assigned(pSD) then
    Exit;
//    raise Exception.Create('Error allocation memory for Security Descriptor');
  if not InitializeSecurityDescriptor(pSD, SECURITY_DESCRIPTOR_REVISION) then
  begin
    LocalFree(NativeUInt(pSD));
    Exit;
//    RaiseLastOSError;
  end;
  // ƒобавить NULL ACL к дескриптору безопасности
  if not SetSecurityDescriptorDacl(pSD, True, nil, False) then
  begin
    LocalFree(NativeUInt(pSD));
    Exit;
//    RaiseLastOSError;
  end;
  Result.nLength := SizeOf(Result);
  Result.lpSecurityDescriptor := pSD;
  Result.bInheritHandle := Inheritable;
end;

initialization
  Randomize;

end.
