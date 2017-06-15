program elevate;

uses
  Windows,
  utils in 'utils.pas';

{$R *.res}

var
  command, sourceFile, varName, value, debugFile, varNewName: string;
  hk: HKEY;
  pipe: THandle;
  code: LONG;
  br: Cardinal;
  debugEnabled: Boolean;
  valueLen: Integer;
  typ: DWORD;
  paramCount: Integer;

  procedure readParamCount;
  var
    cmd: TCommandRec;
  begin
    ReadFile(pipe, cmd, SizeOf(cmd), br, nil);
    paramCount := cmd.ByteCount;
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
  debugFile := ChangeFileExt(GetModuleName(HInstance), '.debug');
  debugEnabled := FileExists(debugFile);

  if (debugEnabled) then
    MessageBox(0, GetCommandLine, 'Command line', 0);

  pipe :=
    CreateFile(
      '\\.\pipe\EnvVarsElevateGate', GENERIC_WRITE, 0, nil, OPEN_EXISTING, 0, 0);
  if (pipe <> INVALID_HANDLE_VALUE) and (pipe <> 0) then
  begin
    command := ParamStr(1);
    varName := ParamStr(2);
    sourceFile := ParamStr(3);
    varNewName := sourceFile;
    if (varName = '') or (sourceFile = '') then
    begin
      MessageBox(0, 'Invalid command', nil, 0);
      ExitCode := ERROR_INVALID_PARAMETER;
    end
    else
    begin
      ExitCode :=
        RegOpenKeyEx(
          HKEY_LOCAL_MACHINE,
          'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
//          HKEY_CURRENT_USER,
//          'Environment',
          0, KEY_READ or KEY_WRITE, hk);
      if (ExitCode = ERROR_SUCCESS) then
      begin
        if (command = '1') then
        begin
          ExitCode := LoadFileData(sourceFile, value);
          if debugEnabled then
            MessageBox(0, PChar('Name: ' + varName + ' Value: ' + value), 'Loaded data', 0);
          if (ExitCode = ERROR_SUCCESS) then
          begin
            if (value <> '') then
            begin
              ExitCode :=
                RegSetValueEx(
                  hk, PChar(varName), 0,
                  IIF(Pos('%', value) > 0, REG_EXPAND_SZ, REG_SZ),
                  PChar(value), (Length(value) + 1) * SizeOf(Char));
            end
            else
            begin
              ExitCode := RegDeleteValue(hk, PChar(varName));
            end;
          end;
        end
        else if (command = '2') then
        begin
          if debugEnabled then
            MessageBox(0, PChar('From: ' + varName + ' To: ' + varNewName), 'Rename', 0);
          ExitCode := RegQueryValueEx(hk, PChar(varName), nil, @typ, nil, @valueLen);
          if (ExitCode = ERROR_SUCCESS) then
          begin
            if (typ in [REG_SZ, REG_EXPAND_SZ]) then
            begin
              SetLength(value, valueLen div 2);
              RegQueryValueEx(hk, PChar(value), nil, nil, @value[1], @valueLen);
              value := Trim(value);
              ExitCode := RegDeleteValue(hk, PChar(varName));
              if (ExitCode = ERROR_SUCCESS) then
              begin
                ExitCode :=
                  RegSetValueEx(
                    hk, PChar(varNewName), 0,
                    IIF(Pos('%', value) > 0, REG_EXPAND_SZ, REG_SZ),
                    PChar(value), (Length(value) + 1) * SizeOf(Char));
              end;
            end;
          end
        end
        else
          ExitCode := ERROR_INVALID_PARAMETER;

        if (debugEnabled) then
        begin
          if (ExitCode <> 0) then
            MessageBox(0, PChar('Error: ' + IntToStr(ExitCode)), nil, 0)
          else
            MessageBox(0, 'Success!', nil, 0);
        end;

        RegCloseKey(hk);
      end;
    end;
    code := ExitCode;
    WriteFile(pipe, code, SizeOf(code), br, nil);
    CloseHandle(pipe);
  end
  else
  begin
    MessageBox(0, 'Gate not found', nil, 0);
    ExitCode := ERROR_FILE_NOT_FOUND;
  end;
end.
