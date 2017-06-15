library envvars;
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) FIELDS([]) PROPERTIES([])}

{$I 'libsuffix.inc'}
{$I 'resources.inc'}
{$R *.res}



{$R 'edit_dialog.res' 'edit_dialog.rc'}
{$R 'properties_dialog.res' 'properties_dialog.rc'}

uses
  Windows,
  fsplugin in 'fsplugin.pas',
  env in 'env.pas',
  utils in 'utils.pas',
  pipes in 'pipes.pas';

var
  PluginNumber: NativeInt;
  ProgressProc: TProgressProcW;
  LogProc: TLogProcW;
  RequestProc: TRequestProcW;


function CheckFlag(N: NativeInt; Flag: NativeUInt): Boolean; inline;
begin
  Result := (N and Flag = Flag);
end;



function FsFindFirstW(APath: PChar; var AFindData: TWin32FindDataW): TEnumerator; stdcall;
begin
  try
    Result := TEnumerator.Create(APath);
    if (Result.HasStrings) then
    begin
      AFindData := Result.Next;
    end
    else
    begin
      Result.Free;
      Result := Pointer(INVALID_HANDLE_VALUE);
    end;
  except
    Result := nil;
  end;
end;



function FsFindNextW(AHandle: TEnumerator; var AFindData: TWin32FindDataW): BOOL; stdcall;
begin
  try
    AFindData := AHandle.Next;
    Result := (AFindData.cFileName[0] <> #0);
  except
    FillChar(AFindData, SizeOf(AFindData), 0);
    Result := False;
  end;
end;



function FsFindClose(AHandle: TEnumerator): NativeInt; stdcall;
begin
  try
    AHandle.Free;
  except
  end;
  Result := 0;
end;



function FsInitW(
  APluginNr: NativeInt; AProgressProcW: TProgressProcW; ALogProcW:
  TLogProcW; ARequestProcW: TRequestProcW): NativeInt; stdcall;
begin
  PluginNumber := APluginNr;
  ProgressProc := AProgressProcW;
  LogProc := ALogProcW;
  RequestProc := ARequestProcW;
  Result := 0;
end;



procedure FsGetDefRootName(ADefRootName: PAnsiChar; AMaxLen: NativeInt); stdcall;
var
  s: AnsiString;
begin
  s := 'Environment Variables Ex'#0;
  Move(s[1], ADefRootName^, Length(s));
end;



function FsInit(PluginNr:integer;pProgressProc:tProgressProc;pLogProc:tLogProc;
                pRequestProc:tRequestProc):integer; stdcall;
begin
  Result := 1;
end;



function FsFindFirst(path :pchar;var FindData:tWIN32FINDDATA):thandle; stdcall;
begin
  Result := ERROR_INVALID_FUNCTION;
end;



function FsFindNext(Hdl:thandle;var FindData:tWIN32FINDDATA):bool; stdcall;
begin
  Result := False;
end;



function FsExtractCustomIconW(
  ARemoteName: PChar; ExtractFlags: NativeInt; var ATheIcon: HICON): NativeInt; stdcall;
var
  loader: TLoader;
begin
  Result := FS_ICON_USEDEFAULT;
  try
    loader := TLoader.Create(ARemoteName, False);
    try
      if (not loader.IsDir) then
      begin
        if (loader.IsNewVar) then
          ATheIcon := LoadIcon(HInstance, MakeIntResource(103))
        else
          ATheIcon := LoadIcon(HInstance, MakeIntResource(104));
        Result := FS_ICON_EXTRACTED_DESTROY;
      end
      else if (not loader.IsParentDirStub) then
      begin
//        ATheIcon := LoadIcon(HInstance, MakeIntResource(105));
//        Result := FS_ICON_EXTRACTED_DESTROY;
      end;
    finally
      loader.Free;
    end;
  except
    ATheIcon := 0;
    Result := FS_ICON_USEDEFAULT;
  end;
end;



function FsGetFileW(
  ARemoteName, ALocalName: PChar; ACopyFlags: NativeInt;
  ARemoteInfo: PRemoteInfo): NativeInt; stdcall;

var
  loader: TLoader;

  function DoCopy: Integer;
  var
    f: THandle;
    data: string;
    bw: Cardinal;
  begin
    if (not loader.IsNewVar) then
    begin
      f :=
        CreateFile(
          ALocalName, GENERIC_READ or GENERIC_WRITE, FILE_SHARE_READ, nil,
          CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
      if (f <> INVALID_HANDLE_VALUE) then
      begin
        try
          data := #$FEFF + {loader.Name + '=' +} loader.Value;
          WriteFile(f, data[1], Length(data) * SizeOf(Char), bw, nil);
          Result := FS_FILE_OK;
        finally
          CloseHandle(f);
        end;
      end
      else
        Result := FS_FILE_READERROR;
    end
    else
      Result := FS_FILE_NOTFOUND;
  end;

begin
  try
    loader := TLoader.Create(ARemoteName, True);
    try
      if
        (not loader.IsDir) and (loader.VarKind <> vkNone) and
        (not loader.IsNewVar)
      then
      begin
        if FileExists(ALocalName) then
        begin
          if CheckFlag(ACopyFlags, FS_COPYFLAGS_OVERWRITE) then
          begin
            Result := DoCopy;
            if (Result = FS_FILE_OK) and CheckFlag(ACopyFlags, FS_COPYFLAGS_MOVE) then
              Result := IIF(loader.Delete, Result, FS_FILE_WRITEERROR);
          end
          else if (ACopyFlags = 0) or CheckFlag(ACopyFlags, FS_COPYFLAGS_MOVE) then
            Result := FS_FILE_EXISTS
          else
            Result := FS_FILE_NOTSUPPORTED;
        end
        else
        begin
          Result := DoCopy;
          if (Result = FS_FILE_OK) and CheckFlag(ACopyFlags, FS_COPYFLAGS_MOVE) then
            Result := IIF(loader.Delete, Result, FS_FILE_WRITEERROR);
        end;
      end
      else
        Result := FS_FILE_NOTSUPPORTED;
    finally
      loader.Free;
    end;
  except
    Result := FS_FILE_READERROR;
  end;
end;



function FsPutFileW(
  ALocalName, ARemoteName: PChar; ACopyFlags: NativeInt): NativeInt; stdcall;
var
  value: string;
  loader: TLoader;

  function DoCopy: Integer;
  begin
    if (loader.SetValue(value)) then
    begin
      Result := FS_FILE_OK
    end
    else
    begin
      if (loader.LastError = ERROR_ACCESS_DENIED) then
      begin
        Result :=
          IIF(
            SetVarAsAdmin(loader.Name, ALocalName, '', '') = ERROR_SUCCESS,
            FS_FILE_OK,
            FS_FILE_WRITEERROR);
        if (Result = FS_FILE_OK) then
          loader.PopulateChanges;
      end
      else
        Result := FS_FILE_WRITEERROR;
    end;
  end;

begin
  try
    if (FileExists(ALocalName)) then
    begin
      if (LoadFileData(string(ALocalName), value) = ERROR_SUCCESS) then
      begin
        loader := TLoader.Create(ARemoteName, true);
        try
          if
            (not loader.IsDir) and (loader.VarKind <> vkNone) and
            (not loader.IsNewVar)
          then
          begin
            if loader.IsVarExists then
            begin
              if CheckFlag(ACopyFlags, FS_COPYFLAGS_OVERWRITE) then
              begin
                Result := DoCopy;
                if (Result = FS_FILE_OK) and CheckFlag(ACopyFlags, FS_COPYFLAGS_MOVE) then
                  Result := IIF(DeleteFile(ALocalName), Result, FS_FILE_WRITEERROR);
              end
              else if (ACopyFlags = 0) or CheckFlag(ACopyFlags, FS_COPYFLAGS_EXISTS_SAMECASE) or CheckFlag(ACopyFlags, FS_COPYFLAGS_EXISTS_DIFFERENTCASE) then
                Result := FS_FILE_EXISTS
              else
                Result := FS_FILE_NOTSUPPORTED;
            end
            else
            begin
              Result := DoCopy;
              if (Result = FS_FILE_OK) and CheckFlag(ACopyFlags, FS_COPYFLAGS_MOVE) then
                Result := IIF(DeleteFile(ALocalName), Result, FS_FILE_WRITEERROR);
            end;
          end
          else
            Result := FS_FILE_NOTSUPPORTED;
        finally
          loader.Free;
        end;
      end
      else
        Result := FS_FILE_READERROR;
    end
    else
    begin
      Result := FS_FILE_NOTFOUND;
    end;
  except
    Result := FS_FILE_READERROR;
  end;
end;



function FsRenMovFileW(
  AOldName, ANewName: PChar; AMove, AOverWrite: BOOL;
  ARemoteInfo: PRemoteInfo): NativeInt; stdcall;
var
  loaderOld, loaderNew: TLoader;
begin
  try
    loaderOld := TLoader.Create(AOldName, True);
    try
      if
        (not loaderOld.IsDir) and (loaderOld.VarKind <> vkNone) and
        (not loaderOld.IsNewVar)
      then
      begin
        loaderNew := TLoader.Create(ANewName, True);
        try
          if
            (not loaderNew.IsDir) and (loaderNew.VarKind <> vkNone) and
            (not loaderNew.IsNewVar)
          then
          begin
            if (not AOverWrite) and (loaderNew.IsVarExists) then
            begin
              Result := FS_FILE_EXISTS;
            end
            else
            begin
              if (loaderNew.SetValue(loaderOld.Value)) then
              begin
                if (loaderOld.Delete) then
                begin
                  Result := FS_FILE_OK;
                end
                else
                begin
                  if (loaderOld.LastError = ERROR_ACCESS_DENIED) then
                  begin
                    Result :=
                      IIF(
                        SetVarAsAdmin(loaderOld.Name, '', '', '') = ERROR_SUCCESS,
                        FS_FILE_OK,
                        FS_FILE_WRITEERROR);
                    if (Result = FS_FILE_OK) then
                      loaderOld.PopulateChanges;
                  end
                  else
                    Result := FS_FILE_WRITEERROR;
                end;
              end
              else
              begin
                if (loaderNew.LastError = ERROR_ACCESS_DENIED) then
                begin
                  Result :=
                    IIF(
                      SetVarAsAdmin(loaderOld.Name, '', '', loaderNew.Name) = ERROR_SUCCESS,
                      FS_FILE_OK,
                      FS_FILE_WRITEERROR);
                  if (Result = FS_FILE_OK) then
                    loaderOld.PopulateChanges;
                end
                else
                  Result := FS_FILE_WRITEERROR;
              end;
            end;
          end
          else
            Result := FS_FILE_WRITEERROR;
        finally
          loaderNew.Free;
        end;
      end
      else
        Result := FS_FILE_WRITEERROR;
    finally
      loaderOld.Free;
    end;
  except
    Result := FS_FILE_WRITEERROR;
  end;
end;



function FsDeleteFileW(ARemoteName: PChar): BOOL; stdcall;
var
  loader: TLoader;
begin
  try
    loader := TLoader.Create(ARemoteName, False);
    try
      if not loader.IsDir then
      begin
        Result := loader.Delete;
        if (not Result) and (loader.LastError = ERROR_ACCESS_DENIED) then
        begin
          Result :=
            (SetVarAsAdmin(loader.Name, '', '', '') = ERROR_SUCCESS);
          if (Result) then
            loader.PopulateChanges;
        end;
      end
      else
        Result := False;
    finally
      loader.Free;
    end;
  except
    Result := False;
  end;
end;



function FsExecuteFileW(AMainWin: THandle; ARemoteName, AVerb: PChar): NativeInt; stdcall;
var
  loader: TLoader;
  dialog: TEditDialog;
  props: TPropDialog;
begin
  if (AVerb = 'open') then
  begin
    try
      loader := TLoader.Create(ARemoteName, True);
      try
        Result := FS_EXEC_OK;
        dialog := TEditDialog.Create(
          IIF(loader.IsNewVar, '', loader.Name), loader.Value, AMainWin);
        try
          if (dialog.OK) and (dialog.Name <> '') then
          begin
            if (loader.IsNewVar) then
              loader.Name := dialog.Name;
            if not (loader.SetValue(dialog.Value)) then
            begin
              if (loader.LastError = ERROR_ACCESS_DENIED) then
              begin
                Result :=
                  IIF(
                    SetVarAsAdmin(loader.Name, '', dialog.Value, '') = ERROR_SUCCESS,
                    FS_EXEC_OK,
                    FS_EXEC_ERROR);
                if (Result = FS_EXEC_OK) then
                  loader.PopulateChanges;
              end
              else
                Result := FS_EXEC_ERROR;
            end;
          end;
        finally
          dialog.Free;
        end;
      finally
        loader.Free;
      end;
    except
      Result := FS_EXEC_ERROR;
    end;
  end
  else if (ARemoteName = '\') and (AVerb = 'properties') then
  begin
    try
      Result := FS_EXEC_OK;
      props := TPropDialog.Create(AMainWin);
      try
        //
      finally
        props.Free;
      end;
    except
      Result := FS_EXEC_ERROR;
    end;
  end
  else
    Result := FS_EXEC_OK;
end;



function FsContentGetDefaultViewW(
  AViewContents, AViewHeaders, AViewWidths, AViewOptions: PChar;
  AMaxLen: NativeInt): BOOL; stdcall;
const
  SViewContents: string = '[=<fs>.Value]';
  SViewHeaders: string = 'Value';
  SViewWidths: string = '90,40,100';
  SViewOptions: string = 'auto-adjust-width';

  procedure copyValue(const S: string; Param: PChar);
  var
    l: Integer;
  begin
    l := Min(AMaxLen, Length(S));
    Move(PChar(S)^, Param^, l * SizeOf(Char));
    Param[l] := #0;
  end;
begin
  copyValue(SViewContents, AViewContents);
  copyValue(SViewHeaders, AViewHeaders);
  copyValue(SViewWidths, AViewWidths);
//  copyValue(SViewOptions, AViewOptions);
  AViewOptions[0] := #0;

  Result := True;
end;



function FsContentGetSupportedField(
  AFieldIndex: NativeInt; AFieldName: PAnsiChar; AUnits: PAnsiChar;
  AMaxLen: NativeInt): NativeInt; stdcall;

  procedure copyValue(const S: AnsiString; Param: PAnsiChar);
  var
    l: Integer;
  begin
    l := Min(AMaxLen, Length(S));
    Move(PAnsiChar(S)^, Param^, l * SizeOf(AnsiChar));
    Param[l] := #0;
  end;

begin
  case AFieldIndex of
    0:
      begin
        copyValue('Value', AFieldName);
        Result := ft_stringw;
      end;
    else
      Result := ft_nomorefields;
  end;
end;



function FsContentGetValueW(
  AFileName: PChar; AFieldIndex, AUnitIndex: NativeInt; AFieldValue: PByte;
  AMaxLen, AFlags: NativeInt): NativeInt; stdcall;
var
  loader: TLoader;
  l: Integer;
begin
  Result := ft_fieldempty;
  try
    loader := TLoader.Create(AFileName, True);
    try
      PDWORD64(AFieldValue)^ := 0;
      if (not loader.IsDir) and (not loader.IsNewVar) then
      begin
        case AFieldIndex of
          0:
            begin
              l := Min(Length(loader.Value), AMaxLen div 2);
              Move(PChar(loader.Value)^, AFieldValue^, l * SizeOf(Char));
              PChar(AFieldValue)[l] := #0;
              Result := ft_stringw;
            end;
        end;
      end;
    finally
      loader.Free;
    end;
  except
    Result := ft_fieldempty;
  end;
end;




exports
  FsInitW,
  FsInit,
  FsGetDefRootName,
  FsFindFirstW,
  FsFindFirst,
  FsFindNextW,
  FsFindNext,
  FsFindClose,
  FsExtractCustomIconW,
  FsGetFileW,
  FsPutFileW,
  FsDeleteFileW,
  FsExecuteFileW,
  FsContentGetDefaultViewW,
  FsContentGetSupportedField,
  FsContentGetValueW,
  FsRenMovFileW;

begin
end.
