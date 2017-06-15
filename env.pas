unit env;

interface
uses
  Windows{$IFDEF DEBUG}{, SysUtils}{$ENDIF}, utils, Messages;


type

  TVarKind = (vkProcess, vkVolatile, vkUser, vkSystem, vkNone);

  TLocalization = class
  private var
    FLangLoaded: Boolean;
    FLangName: string;
    FName: string;
    FNewVar: string;
    FChangeVariable: string;
    FSystem: string;
    FValue: string;
    FProcess: string;
    FUser: string;
    FProperties: string;
    FLanguage: string;
    FVolatile: string;
  public
    constructor Create;

    property LangName: string read FLangName;
    property System: string read FSystem write FSystem;
    property User: string read FUser write FUser;
    property Volatile: string read FVolatile write FVolatile;
    property Process: string read FProcess write FProcess;
    property ChangeVariable: string read FChangeVariable write FChangeVariable;
    property Name: string read FName write FName;
    property Value: string read FValue write FValue;
    property NewVar: string read FNewVar write FNewVar;
    property Properties: string read FProperties write FProperties;
    property Language: string read FLanguage write FLanguage;

    function GetVarKind(const AKind: string): TVarKind;
    function IsNewVar(const AName: string): Boolean;
    function GetLangList: TStringArray;
    procedure SetLang(const ALang: string);
  end;


  TEnumerator = class
  private
    FLocalization: TLocalization;
    FPosition: Integer;
    FStrings: string;
    function LoadExternalStrings(AType: TVarKind): string;
    function LoadEnvStrings: string;
    function GetHasStrings: Boolean;
    function GetStringItem: string;
  public
    constructor Create(const APath: string);
    destructor Destroy; override;

    property HasStrings: Boolean read GetHasStrings;

    function Next: TWin32FindDataW;
  end;


  TLoader = class
  private
    FLocalization: TLocalization;
    FIsDir: Boolean;
    FName: string;
    FValue: string;
    FVarKind: TVarKind;
    FLastError: NativeInt;
    function LoadExternalString: string;
    function LoadEnvString: string;
    function RemoveExternalString: Boolean;
    function RemoveEnvString: Boolean;
    function SetExternalString: Boolean;
    function SetEnvString: Boolean;
    function GetIsNewVar: Boolean;
    function GetIsParentDirStub: Boolean;
    function GetIsVarExists: Boolean;
  public
    constructor Create(const APath: string; ADoLoadData: Boolean);
    destructor Destroy; override;

    property VarKind: TVarKind read FVarKind;
    property Name: string read FName write FName;
    property Value: string read FValue;
    property IsDir: Boolean read FIsDir;
    property IsNewVar: Boolean read GetIsNewVar;
    property IsParentDirStub: Boolean read GetIsParentDirStub;
    property IsVarExists: Boolean read GetIsVarExists;
    property LastError: NativeInt read FLastError;

    function Delete: Boolean;
    function SetValue(const ANewValue: string): Boolean;
    procedure PopulateChanges;
  end;


  TEditDialog = class
  private
    FName: string;
    FValue: string;
    FOK: Boolean;
    FLocalization: TLocalization;
    procedure DialogProc(
      AWnd: HWND; AMsg: NativeUInt; WParam, LParam: NativeInt);
  public
    constructor Create(const AName, AValue: string; AParent: HWND);
    destructor Destroy; override;

    property Name: string read FName;
    property Value: string read FValue;
    property OK: Boolean read FOK;
  end;

  TPropDialog = class
  private
    FOK: Boolean;
    FLocalization: TLocalization;
    FLang: string;
    FLangList: TStringArray;
    procedure DialogProc(
      AWnd: HWND; AMsg: NativeUInt; WParam, LParam: NativeInt);
  public
    constructor Create(AParent: HWND);
    destructor Destroy; override;

    property Lang: string read FLang;
    property OK: Boolean read FOK;
  end;


implementation


function OpenEnvKey(AType: TVarKind; AWrite: Boolean; var AError: NativeInt): HKEY;
var
  rootKey: HKEY;
  key: string;
begin
  if (AType = vkSystem) then
  begin
    rootKey := HKEY_LOCAL_MACHINE;
    key := 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment';
  end
  else if (AType = vkVolatile) then
  begin
    rootKey := HKEY_CURRENT_USER;
    key := 'Volatile Environment';
  end
  else if (AType = vkUser) then
  begin
    rootKey := HKEY_CURRENT_USER;
    key := 'Environment';
  end
  else
    rootKey := 0;
  if (rootKey <> 0) then
  begin
    AError :=
      RegOpenKeyEx(
        rootKey, PChar(key), 0, IIF(AWrite, KEY_WRITE, KEY_READ), Result);
    if (AError <> ERROR_SUCCESS) then
      Result := 0;
  end
  else
    Result := 0;
end;


{ TEnumerator }


constructor TEnumerator.Create(const APath: string);
var
  s, kind: string;
  vk: TVarKind;
begin
  inherited Create;
  FLocalization := TLocalization.Create;

  // Load localization here
  if (APath = '\') then
  begin
    FStrings :=
      FLocalization.System + #0 + FLocalization.User + #0 +
      FLocalization.Volatile + #0 + FLocalization.Process;
    FPosition := 1;
  end
  else
  begin
    s := APath;
    kind := ExtractPathItem(s);
    vk := FLocalization.GetVarKind(kind);
    case vk of
      vkProcess:
        begin
          FStrings := LoadEnvStrings;
          FPosition := 1;
        end;

      vkUser,
      vkSystem,
      vkVolatile:
        begin
          FStrings := LoadExternalStrings(vk);
          FPosition := 1;
        end;
    end
  end;
end;

destructor TEnumerator.Destroy;
begin
  FLocalization.Free;
  inherited;
end;

function TEnumerator.GetHasStrings: Boolean;
begin
  Result := (FStrings <> '');
end;

function TEnumerator.GetStringItem: string;
var
  i, l: Integer;
begin
  l := Length(FStrings);
  for i := FPosition to l do
  begin
    if (FStrings[i] = #0) then
    begin
      Result := Copy(FStrings, FPosition, i - FPosition);
      FPosition := i + 1;
      break;
    end
    else if (i = l) then
    begin
      Result := Copy(FStrings, FPosition, i - FPosition + 1);
      FPosition := i + 1;
    end;
  end;
end;

function TEnumerator.LoadEnvStrings: string;
var
  item, name, value: string;
  block, ptr, last: PChar;
  delPos: Integer;
begin
  Result := FLocalization.NewVar + #9;
  block := GetEnvironmentStrings;
  try
    ptr := block;
    last := ptr;
    repeat
      Inc(ptr);
      if (ptr^ = #0) then
      begin
        item := last;
        delPos := Pos('=', item);
        if (delPos > 1) then
        begin
          name := Copy(item, 1, delPos - 1);
          value := Copy(item, delPos + 1, Length(item));
          if (Result <> '') then
            Result := Result + #0;
          Result := Result + name + #9 + value;
        end;
        last := ptr;
        Inc(last);
      end;
    until (ptr[0] = #0) and (ptr[1] = #0);
  finally
    FreeEnvironmentStrings(block);
  end;
end;

function TEnumerator.LoadExternalStrings(AType: TVarKind): string;
var
  hk: HKEY;
  value, name: string;
  valuesCount, maxNameLen, maxValueLen, nameLen, valueLen: DWORD;
  i: Integer;
  err: NativeInt;
begin
  Result := FLocalization.NewVar + #9;
  hk := OpenEnvKey(AType, False, err);
  if (hk <> 0) then
  begin
    try
      if (
        (RegQueryInfoKey(
          hk, nil, nil, nil, nil, nil, nil, @valuesCount, @maxNameLen,
          @maxValueLen, nil, nil) = ERROR_SUCCESS) and
          (valuesCount > 0))
      then
      begin
        Inc(maxNameLen);
        SetLength(name, maxNameLen);
        SetLength(value, maxValueLen div 2);
        for i := 0 to valuesCount - 1 do
        begin
          nameLen := maxNameLen;
          valueLen := maxValueLen;
          if (
            RegEnumValue(
              hk, i, @name[1], nameLen, nil, nil, @value[1],
              @valueLen) = ERROR_SUCCESS)
          then
          begin
            if (Result <> '') then
              Result := Result + #0;

            Result :=
              Result + Trim(Copy(name, 1, nameLen)) + #9 +
              Trim(Copy(value, 1, valueLen div 2));
          end;
        end;
      end;
    finally
      RegCloseKey(hk);
    end;
  end;
end;

function TEnumerator.Next: TWin32FindDataW;
var
  delPos: Integer;
  s, name, value: string;
begin
  FillChar(Result, SizeOf(Result), 0);
  s := GetStringItem;
  if (s <> '') then
  begin
    delPos := Pos(#9, s);
    if (delPos > 0) then
    begin
      name := Copy(s, 1, delPos - 1);
      value := Copy(s, delPos + 1, Length(s));
      Result.nFileSizeLow := Length(value);
      s := name;
      Move(s[1], Result.cFileName, Min(Length(s), (SizeOf(Result.cFileName) div 2) - 1) * SizeOf(Char));
    end
    else
    begin
      Result.dwFileAttributes := FILE_ATTRIBUTE_DIRECTORY;
      Move(s[1], Result.cFileName, Min(Length(s), (SizeOf(Result.cFileName) div 2) - 1) * SizeOf(Char));
    end;
  end;
end;


{ TLoader }


constructor TLoader.Create(const APath: string; ADoLoadData: Boolean);
var
  s, kind: string;
  delPos: Integer;
begin
  inherited Create;
  FLocalization := TLocalization.Create;

  FIsDir := APath[Length(APath)] = '\';

  if (not IsDir) then
  begin
    s := APath;
    kind := ExtractPathItem(s);
    FVarKind := FLocalization.GetVarKind(kind);
    delPos := Pos(' = ', s);
    if (delPos > 0) then
      s := Copy(s, 1, delPos - 1);

    FName := s;

    if (ADoLoadData) then
    begin
      case VarKind of
        vkProcess: FValue := LoadEnvString;
        vkUser,
        vkSystem,
        vkVolatile:
          FValue := LoadExternalString;
      end;
    end;
  end
  else
  begin
    s := Copy(APath, 1, Length(APath) - 1);
    kind := ExtractPathItem(s);
    if (kind = '..') then
    begin
      FVarKind := vkNone;
      FName := kind;
    end
    else
    begin
      FVarKind := FLocalization.GetVarKind(kind);
      FName := s;
    end;
  end;
end;

function TLoader.Delete: Boolean;
begin
  case VarKind of
    vkProcess: Result := RemoveEnvString;
    vkUser: Result := RemoveExternalString;
    vkSystem: Result := RemoveExternalString;
    else
      Result := False;
  end;
end;

destructor TLoader.Destroy;
begin
  FLocalization.Free;
  inherited;
end;

function TLoader.GetIsNewVar: Boolean;
begin
  Result := FLocalization.IsNewVar(Name);
end;

function TLoader.GetIsParentDirStub: Boolean;
begin
  Result := (Name = '..');
end;

function TLoader.GetIsVarExists: Boolean;
begin
  Result := not IsNewVar and not IsDir and (Value <> '');
end;

function TLoader.LoadEnvString: string;
begin
  Setlength(Result, Windows.GetEnvironmentVariable(PChar(FName), nil, 0));
  Windows.GetEnvironmentVariable(PChar(FName), @Result[1], Length(Result));
  Result := Trim(Result);
end;

function TLoader.LoadExternalString: string;
var
  hk: HKEY;
  valueLen: Integer;
  typ: DWORD;
//  err: NativeInt;
begin
  hk := OpenEnvKey(VarKind, False, FLastError);
  if (hk <> 0) then
  begin
    try
      if (RegQueryValueEx(hk, PChar(FName), nil, @typ, nil, @valueLen) = ERROR_SUCCESS) then
      begin
        if (typ in [REG_SZ, REG_EXPAND_SZ]) then
        begin
          SetLength(Result, valueLen div 2);
          RegQueryValueEx(hk, PChar(FName), nil, nil, @Result[1], @valueLen);
          Result := Trim(Result);
        end;
      end
    finally
      RegCloseKey(hk);
    end;
  end;
end;

procedure TLoader.PopulateChanges;
begin
  SendMessageTimeout(
    HWND_BROADCAST, WM_SETTINGCHANGE, 0, Integer(PChar('Environment')),
    SMTO_NORMAL, 1000, nil);
end;

function TLoader.RemoveEnvString: Boolean;
begin
  Result := Windows.SetEnvironmentVariable(PChar(FName), nil);
end;

function TLoader.RemoveExternalString: Boolean;
var
  hk: HKEY;
begin
  hk := OpenEnvKey(VarKind, True, FLastError);
  if (hk <> 0) then
  begin
    try
      Result := (RegDeleteValue(hk, PChar(FName)) = ERROR_SUCCESS);
      if (Result) then
        PopulateChanges;
    finally
      RegCloseKey(hk);
    end;
  end
  else
  begin
    Result := False;
  end;
end;

function TLoader.SetEnvString: Boolean;
begin
  Result := Windows.SetEnvironmentVariable(PChar(FName), PChar(FValue));
end;

function TLoader.SetExternalString: Boolean;
var
  hk: HKEY;
  s: string;
begin
  hk := OpenEnvKey(VarKind, True, FLastError);
  if (hk <> 0) then
  begin
    try
      s := FValue + #0;
      Result :=
        (RegSetValueEx(
          hk, PChar(FName), 0, IIF(Pos('%', FValue) > 0, REG_EXPAND_SZ, REG_SZ),
          PChar(s), (Length(s) + 1) * SizeOf(Char)) = ERROR_SUCCESS);
      if (Result) then
        PopulateChanges;
    finally
      RegCloseKey(hk);
    end;
  end
  else
  begin
    // try to elevate
    {if (err = ERROR_ACCESS_DENIED) then
      Result := (SetVarAsAdmin(AName, '', AValue) = ERROR_SUCCESS)
    else}
      Result := False;
  end;
end;

function TLoader.SetValue(const ANewValue: string): Boolean;
begin
  if (ANewValue <> '') then
  begin
    FValue := ANewValue;
    case VarKind of
      vkProcess: Result := SetEnvString;
      vkUser: Result := SetExternalString;
      vkSystem: Result := SetExternalString;
      else
        Result := False;
    end;
  end
  else
    Result := Delete;
end;


{ TLocalization }


constructor TLocalization.Create;

  function GetValue(const ASection, AName, ADefault, AIni: string): string;
  begin
    Setlength(Result, 4096);
    SetLength(
      Result,
      GetPrivateProfileString(
        PChar(ASection), PChar(AName), PChar(''), @Result[1], 4096, PChar(AIni)));
  end;

  procedure GetLangData(AIni: string);
  var
    f: THandle;
    sz, linePos, delPos: Integer;
    br: Cardinal;
    data, line, name, value: string;

    function GetLine: Boolean;
    var
      i: Integer;
    begin
      for i := linePos to length(data) do
      begin
        if (data[i] = #13) then
        begin
          line := Trim(Copy(data, linePos, i - linePos));
          linePos := i + 1;
          if (line <> '') then
          begin
            if (line[1] <> ';') then
            begin
              Break;
            end;
            line := '';
          end;
        end;
      end;
      Result := (line <> '');
    end;

  begin
    f :=
      CreateFile(
        PChar(AIni), GENERIC_READ, FILE_SHARE_READ, nil, OPEN_EXISTING, 0, 0);
    if (f <> INVALID_HANDLE_VALUE) then
    begin
      try
        sz := Min(GetFileSize(f, nil), 4096);
        SetLength(data, sz);
        ReadFile(f, data[1], sz, br, nil);
      finally
        CloseHandle(f);
      end;
      linePos := 1;
      while GetLine do
      begin
        delPos := Pos('=', line);
        if (delPos > 0) then
        begin
          name := Trim(Copy(line, 1, delPos - 1));
          value := Trim(Copy(line, delPos + 1, Length(line)));
          if (name = '[New variable]') then
            FNewVar := value
          else if (name = 'Current process (Total Commander)') then
            FProcess := value
          else if (name = 'System') then
            FSystem := value
          else if (name = 'Current user') then
            FUser := value
          else if (name = 'Name:') then
            FName := value
          else if (name = 'Value:') then
            FValue := value
          else if (name = 'Change variable') then
            FChangeVariable := value
          else if (name = 'Language:') then
            FLanguage := value
          else if (name = 'Properties') then
            FProperties := value
          else if (name = 'Volatile') then
            FVolatile := value;
        end;
      end;
    end;
  end;

var
  ini, lang, langFile: string;
begin
  if (not FLangLoaded) then
  begin
    FProcess := 'Current process (Total Commander)';
    FUser := 'Current user';
    FSystem := 'System';
    FVolatile := 'Volatile';
    FChangeVariable := 'Change Variable';
    FName := 'Name:';
    FValue := 'Value:';
    FNewVar := '[New variable]';
    FLanguage := 'Language:';
    FProperties := 'Properties';
    ini := GetModuleName(HInstance);
    ini := ExtractFilePath(ini) + 'envvars.ini';

    if (FileExists(ini)) then
    begin
      lang := GetValue('General', 'Lang', '', ini);
      if (lang <> '') then
      begin
        langFile := ExtractFilePath(ini) + lang;
        if (FileExists(langFile)) then
        begin
          GetLangData(langFile);
          FLangName := ExtractFileName(langFile);
        end;
      end;
    end;
    FLangLoaded := True;
  end;
end;

function TLocalization.GetLangList: TStringArray;
begin
  SetLength(Result, 1);
  Result[0] := 'Default';
  DynArrayAppend(Result, GetFileList(ExtractFilePath(GetModuleName(HInstance)) + 'Lang', '*.lng'));
end;

function TLocalization.GetVarKind(const AKind: string): TVarKind;
begin
  if (AKind = Process) then
    Result := vkProcess
  else if (AKind = System) then
    Result := vkSystem
  else if (AKind = User) then
    Result := vkUser
  else if (AKind = Volatile) then
    Result := vkVolatile
  else
    Result := vkProcess;
end;

function TLocalization.IsNewVar(const AName: string): Boolean;
begin
  Result := (AName = NewVar);
end;


procedure TLocalization.SetLang(const ALang: string);
var
  ini, lang: string;
begin
  ini := GetModuleName(HInstance);
  ini := ExtractFilePath(ini) + 'envvars.ini';

  lang := 'Lang\' + ALang;
  WritePrivateProfileString('General', 'Lang', PChar(lang), PChar(ini));
  FLangLoaded := False;
end;

{ TEditDialog }

const
  EDIT_IDNAME = 1001;
  EDIT_IDVALUE = 1002;
  EDIT_IDNAMELABEL = 1005;
  EDIT_IDVALUELABEL = 1006;

function Edit_DialogProc(AWnd: HWND; AMsg: NativeUInt; WParam, LParam: NativeInt): NativeInt; stdcall;
var
  dlg: TEditDialog;
begin
  case AMsg of
    WM_INITDIALOG:
      begin
        SetWindowLongPtr(AWnd, GWLP_USERDATA, LParam);
        dlg := Pointer(LParam);
        dlg.DialogProc(AWnd, AMsg, WParam, LParam);
      end;

    else
    begin
      dlg := Pointer(GetWindowLongPtr(AWnd, GWLP_USERDATA));
      if (dlg <> nil) then
        dlg.DialogProc(AWnd, AMsg, WParam, LParam);
    end;
  end;
  Result := 0;
end;

constructor TEditDialog.Create(const AName, AValue: string; AParent: HWND);
begin
  inherited Create;
  FLocalization := TLocalization.Create;
  FName := AName;
  FValue := AValue;
  FOK := (DialogBoxParam(
    HInstance, MakeIntResource(101), AParent, @Edit_DialogProc,
    Integer(Self)) = ID_OK);
end;

destructor TEditDialog.Destroy;
begin
  FLocalization.Free;
  inherited;
end;

procedure TEditDialog.DialogProc(
  AWnd: HWND; AMsg: NativeUInt; WParam, LParam: NativeInt);
var
  wnd: HWND;
begin
  case AMsg of
    WM_INITDIALOG:
      begin
        if (Name = '') then
        begin
          wnd := GetDlgItem(AWnd, EDIT_IDNAME);
          SendMessage(wnd, EM_SETREADONLY, 0, 0);
          SetWindowText(AWnd, PChar(FLocalization.NewVar));
        end
        else
        begin
          SetDlgItemText(AWnd, EDIT_IDNAME, PChar(Name));
          SetWindowText(AWnd, PChar(FLocalization.ChangeVariable));
        end;
        SetDlgItemText(AWnd, EDIT_IDVALUE, PChar(Value));
        SetDlgItemText(AWnd, EDIT_IDNAMELABEL, PChar(FLocalization.Name));
        SetDlgItemText(AWnd, EDIT_IDVALUELABEL, PChar(FLocalization.Value));
      end;

    WM_COMMAND:
      begin
        case LOWORD(WParam) of
          ID_OK:
            begin
              wnd := GetDlgItem(AWnd, EDIT_IDVALUE);
              SetLength(FValue, SendMessage(wnd, WM_GETTEXTLENGTH, 0, 0) + 1);
              SendMessage(wnd, WM_GETTEXT, Length(FValue), Integer(@FValue[1]));
              FValue := Trim(FValue);

              wnd := GetDlgItem(AWnd, EDIT_IDNAME);
              SetLength(FName, SendMessage(wnd, WM_GETTEXTLENGTH, 0, 0) + 1);
              SendMessage(wnd, WM_GETTEXT, Length(FName), Integer(@FName[1]));
              FName := Trim(FName);
              EndDialog(AWnd, ID_OK);
            end;

          ID_CANCEL:
            EndDialog(AWnd, ID_CANCEL);
        end;
      end;
  end;
end;


{ TPropDialog }

const
  PROP_IDLANG = 1008;
  PROP_IDLANGLABEL = 1009;

function Prop_DialogProc(AWnd: HWND; AMsg: NativeUInt; WParam, LParam: NativeInt): NativeInt; stdcall;
var
  dlg: TPropDialog;
begin
  case AMsg of
    WM_INITDIALOG:
      begin
        SetWindowLongPtr(AWnd, GWLP_USERDATA, LParam);
        dlg := Pointer(LParam);
        dlg.DialogProc(AWnd, AMsg, WParam, LParam);
      end;

    else
    begin
      dlg := Pointer(GetWindowLongPtr(AWnd, GWLP_USERDATA));
      if (dlg <> nil) then
        dlg.DialogProc(AWnd, AMsg, WParam, LParam);
    end;
  end;
  Result := 0;
end;

constructor TPropDialog.Create(AParent: HWND);
begin
  inherited Create;
  FLocalization := TLocalization.Create;
  FLangList := FLocalization.GetLangList;
  FOK := (DialogBoxParam(
    HInstance, MakeIntResource(106), AParent, @Prop_DialogProc,
    Integer(Self)) = ID_OK);
end;

destructor TPropDialog.Destroy;
begin
  SetLength(FLangList, 0);
  FLocalization.Free;
  inherited;
end;

procedure TPropDialog.DialogProc(
  AWnd: HWND; AMsg: NativeUInt; WParam, LParam: NativeInt);
var
  list: HWND;
  i, n: Integer;
  s: string;
begin
  case AMsg of
    WM_INITDIALOG:
      begin
        list := GetDlgItem(AWnd, PROP_IDLANG);
        SendMessage(list, CB_RESETCONTENT, 0, 0);
        for i := 0 to Length(FLangList) - 1 do
        begin
          s := FLangList[i];
          n := SendMessage(list, CB_ADDSTRING, 0, NativeInt(PChar(s)));
          SendMessage(list, CB_SETITEMDATA, n, i);
          if (FLocalization.LangName = s) or (i = 0) then
            SendMessage(list, CB_SETCURSEL, n, 0);
        end;

        SetWindowText(AWnd, PChar(FLocalization.Properties));
        SetDlgItemText(AWnd, PROP_IDLANGLABEL, PChar(FLocalization.Language));
      end;

    WM_COMMAND:
      begin
        case LOWORD(WParam) of
          ID_OK:
            begin
              list := GetDlgItem(AWnd, PROP_IDLANG);
              n := SendMessage(list, CB_GETITEMDATA, SendMessage(list, CB_GETCURSEL, 0, 0), 0);
              if (n > 0) then
              begin
                FLang := FLangList[n];
                FLocalization.SetLang(FLang);
              end
              else
                FLang := '';
              EndDialog(AWnd, ID_OK);
            end;

          ID_CANCEL:
            EndDialog(AWnd, ID_CANCEL);
        end;
      end;
  end;
end;

end.
