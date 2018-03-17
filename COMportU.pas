unit COMportU;

// Please, don't delete this comment. \\
(*
  Copyright Owner: Yahe
  Copyright Year : 2008-2018

  Unit   : COMportU (platform dependant)
  Version: 1.3c

  Contact E-Mail: hello@yahe.sh
*)
// Please, don't delete this comment. \\

(*
  Description:

  This unit contains an interface to COM ports (/serial ports).
*)

(*
  Change Log:

  [Version 1.3c] (10.11.2008: write release)
  - PurgeBuffers() introduced
  - Read() introduced
  - ReadLine() introduced
  - Write() introduced
  - WriteLine() introduced
  - class TCOMportRead renamed to TCOMportThread
  - all COM port access is now synchronized
  - CloseCOMhandle() now member of TCOMport
  - OpenCOMhandle() now member of TCOMport
  - Synchronized property removed
  - Threaded property introduced
  - Timeout property introduced
  - ClearBufferOnClose property introduced
  - TSetupConnectionEvent changed to handle Timeouts

  [Version 1.2.1c] (10.10.2008: dialog fix release)
  - config-dialog fixed

  [Version 1.2c] (26.08.2008: sender release)
  - all events now provide a Sender parameter

  [Version 1.1c] (19.07.2008: synchronize release)
  - TCOMport is now synchronizable
  - buffer can now be flushed after being opened

  [Version 1.0c] (16.07.2008: initial release)
  - initial source has been written
*)

interface

uses
  Windows,
  SysUtils,
  SyncObjs,
  Classes;

const
  COMportU_CopyrightOwner = 'Yahe';
  COMportU_CopyrightYear  = '2008-2018';
  COMportU_Name           = 'COMport';
  COMportU_ReleaseDate    = '10.11.2008';
  COMportU_ReleaseName    = 'write release';
  COMportU_Version        = '1.3c';

const
  CopyrightOwner = COMportU_CopyrightOwner;
  CopyrightYear  = COMportU_CopyrightYear;
  Name           = COMportU_Name;
  ReleaseDate    = COMportU_ReleaseDate;
  ReleaseName    = COMportU_ReleaseName;
  Version        = COMportU_Version;

type
  TCOMport       = class;
  TCOMportThread = class;

  TReadCharEvent        = procedure (const ASender : TObject; const AChar : Char) of Object;
  TReadLineEvent        = procedure (const ASender : TObject; const ALine : String) of Object;
  TSetupConnectionEvent = procedure (const ASender : TObject; const ACOMhandle : THandle; var AConfig : TCommConfig; var ATimeouts : TCommTimeouts) of Object;

  TCOMport = class(TObject)
  private
  protected
    FClearBufferOnClose : Boolean;
    FClearBufferOnOpen  : Boolean;
    FCOMhandle          : THandle;
    FConfig             : TCommConfig;
    FFlushBufferOnClose : Boolean;
    FLineBuffer         : String;
    FLineEnd            : String;
    FNumber             : Byte;
    FReadOnly           : Boolean;
    FShowSetupDialog    : Boolean;
    FThread             : TCOMportThread;
    FThreaded           : Boolean;
    FTimeouts           : TCommTimeouts;

    FReadLineEvent        : TReadLineEvent;
    FSetupConnectionEvent : TSetupConnectionEvent;

    function CloseCOMhandle : Boolean;
    function OpenCOMhandle(const ACOMname : String) : Boolean;

    procedure CreateThread;
    procedure DestroyThread;

    procedure ReadChar(const ASender : TObject; const AChar : Char);
    procedure SetConfig(const AValue : TCommConfig);
    procedure SetNumber(const AValue : Byte);
    procedure SetReadOnly(const AValue : Boolean);
    procedure SetThreaded(const AValue : Boolean);
    procedure SetTimeouts(const AValue : TCommTimeouts);
  public
    constructor Create(const ANumber : Byte);

    destructor Destroy; override;

    property ClearBufferOnClose : Boolean       read FClearBufferOnClose write FClearBufferOnClose;
    property ClearBufferOnOpen  : Boolean       read FClearBufferOnOpen  write FClearBufferOnOpen;
    property Config             : TCommConfig   read FConfig             write SetConfig;
    property FlushBufferOnClose : Boolean       read FFlushBufferOnClose write FFlushBufferOnClose;
    property LineEnd            : String        read FLineEnd            write FLineEnd;
    property Number             : Byte          read FNumber             write SetNumber;
    property ReadOnly           : Boolean       read FReadOnly           write SetReadOnly;
    property ShowSetupDialog    : Boolean       read FShowSetupDialog    write FShowSetupDialog;
    property Threaded           : Boolean       read FThreaded           write SetThreaded;
    property Timeouts           : TCommTimeouts read FTimeouts           write SetTimeouts;

    property OnReadLine        : TReadLineEvent        read FReadLineEvent        write FReadLineEvent;
    property OnSetupConnection : TSetupConnectionEvent read FSetupConnectionEvent write FSetupConnectionEvent;

    function CloseConnection : Boolean;
    function OpenConnection : Boolean;

    function IsOpen : Boolean;

    function PurgeBuffers(const ASpecification : LongInt = - 1) : Boolean;

    function Read(const ALength : LongWord) : String;
    function ReadLine : String;

    function Write(const AString : String) : LongWord;
    function WriteLine(const AString : String) : Boolean;
  published
  end;

  TCOMportThread = class(TThread)
  private
  protected
    FCOMhandle       : THandle;
    FCriticalSection : TCriticalSection;

    FReadCharEvent : TReadCharEvent;

    procedure Execute; override;
  public
    constructor Create(const ACreateSuspended : Boolean);

    destructor Destroy; override;

    property COMhandle : THandle read FCOMhandle write FCOMhandle;
    
    property OnReadChar : TReadCharEvent read FReadCharEvent write FReadCharEvent;

    procedure EnterSection;
    procedure LeaveSection;
  published
  end;

const
  CCOMportShort = 'COM';
  CCOMportLong  = '\\.\' + CCOMportShort;

implementation

{ TCOMport }

function TCOMport.CloseConnection : Boolean;
begin
  Result := false;

  if IsOpen then
  begin
    if Threaded then
      DestroyThread;

    if (Length(FLineBuffer) > 0) then
    begin
      if FFlushBufferOnClose then
      begin
        if Assigned(FReadLineEvent) then
          FReadLineEvent(Self, FLineBuffer);
      end;
    end;
    FLineBuffer := '';

    if (FClearBufferOnClose) then
      PurgeBuffers;

    Result := CloseCOMhandle;
  end;
end;

constructor TCOMport.Create(const ANumber: Byte);
begin
  inherited Create;

  FClearBufferOnOpen    := true;
  FCOMhandle            := INVALID_HANDLE_VALUE;
  FLineBuffer           := '';
  FLineEnd              := #13#10;
  FNumber               := ANumber;
  FReadOnly             := false;
  FSetupConnectionEvent := nil;
  FShowSetupDialog      := true;
  FThreaded             := false;

  OnReadLine        := nil;
  OnSetupConnection := nil;

  FConfig.dcb.DCBlength := SizeOf(FConfig.dcb);
  FConfig.dwSize        := SizeOf(FConfig);

  FTimeouts.ReadIntervalTimeout         := 1000;
  FTimeouts.ReadTotalTimeoutMultiplier  := 1000;
  FTimeouts.ReadTotalTimeoutConstant    := 0;
  FTimeouts.WriteTotalTimeoutMultiplier := 1000;
  FTimeouts.WriteTotalTimeoutConstant   := 0;
end;

destructor TCOMport.Destroy;
begin
  CloseConnection;

  inherited Destroy;
end;

function TCOMport.OpenConnection : Boolean;
begin
  Result := false;

  if not(IsOpen) then
  begin
    Result := OpenCOMhandle(CCOMportLong + IntToStr(FNumber));
    if Result then
    begin
      FLineBuffer := '';

      if FShowSetupDialog then
        CommConfigDialog(PChar(CCOMportShort + IntToStr(FNumber)), 0, FConfig);

      if Assigned(FSetupConnectionEvent) then
        FSetupConnectionEvent(Self, FCOMhandle, FConfig, FTimeouts);

      SetConfig(FConfig);
      SetTimeouts(FTimeouts);

      if Threaded then
        CreateThread;
    end;
  end;
end;

function TCOMport.CloseCOMhandle : Boolean;
begin
  Result := false;

  if IsOpen then
  begin
    Result := CloseHandle(FCOMhandle);

    FCOMhandle := INVALID_HANDLE_VALUE;
  end;
end;

function TCOMport.OpenCOMhandle(const ACOMname: String): Boolean;
var
  LAccess : LongWord;
  LSize   : LongWord;
begin
  Result := false;

  if not(IsOpen) then
  begin
    LAccess := GENERIC_READ;
    if not(FReadOnly) then
      LAccess := LAccess or GENERIC_WRITE;

    FCOMhandle := CreateFile(PChar(ACOMname),
                             LAccess,
                             0, // port not shareable
                             nil,
                             OPEN_EXISTING, // necessary for serial port
                             FILE_ATTRIBUTE_NORMAL,
                             0);

    if IsOpen then
    begin
      if FClearBufferOnOpen then
        Result := PurgeBuffers;

      if Result then
      begin
        Result := GetCommConfig(FCOMhandle, FConfig, LSize);
        if Result then
          Result := GetCommTimeouts(FCOMhandle, FTimeouts);
      end;

      if not(Result) then
        CloseCOMhandle;
    end;
  end;
end;

procedure TCOMport.ReadChar(const ASender : TObject; const AChar : Char);
var
  LLine : String;
begin
  if Assigned(FReadLineEvent) then
  begin
    if (Length(FLineEnd) > 0) then
    begin
      FLineBuffer := FLineBuffer + AChar;
      while (Pos(FLineEnd, FLineBuffer) > 0) do
      begin
        LLine := Copy(FLineBuffer, 1, Pred(Pos(FLineEnd, FLineBuffer)));
        Delete(FLineBuffer, 1, Pos(FLineEnd, FLineBuffer) + Pred(Length(FLineEnd)));

        FReadLineEvent(Self, LLine);
      end;
    end
    else
      FReadLineEvent(Self, AChar);
  end;
end;

procedure TCOMport.SetConfig(const AValue : TCommConfig);
begin
  FConfig := AValue;

  if IsOpen then
    SetCommConfig(FCOMhandle, FConfig, SizeOf(FConfig));
end;

procedure TCOMport.SetNumber(const AValue : Byte);
begin
  if not(IsOpen) then
    FNumber := AValue;
end;

procedure TCOMport.SetReadOnly(const AValue : Boolean);
begin
  if not(IsOpen) then
    FReadOnly := AValue;
end;

procedure TCOMport.SetThreaded(const AValue : Boolean);
begin
  if (AValue <> FThreaded) then
  begin
    if AValue then
    begin
      if IsOpen then
        CreateThread;
    end
    else
      DestroyThread;

    FThreaded := AValue;
  end;
end;

function TCOMport.Write(const AString : String) : LongWord;
begin
  Result := 0;

  if IsOpen then
  begin
    if Threaded then
      FThread.EnterSection;
    try
      WriteFile(FCOMhandle, AString[1], Length(AString), Result, nil);
    finally
      if Threaded then
        FThread.LeaveSection;
    end;
  end;
end;

function TCOMport.WriteLine(const AString : String) : Boolean;
var
  LLength : LongWord;
begin
  LLength := Length(AString) + Length(FLineEnd);

  Result := (Write(AString + FLineEnd) = LLength);
end;

function TCOMport.Read(const ALength : LongWord) : String;
var
  LBytesRead : LongWord;
  LLength    : LongWord;
begin
  Result := '';

  if (IsOpen and (ALength > 0)) then
  begin
    if Threaded then
      FThread.EnterSection;
    try
      SetLength(Result, ALength);

      ReadFile(FCOMhandle, Result[1], Length(Result), LBytesRead, nil);

      LLength := Length(Result);
      if (LBytesRead < LLength) then
        SetLength(Result, LBytesRead);
    finally
      if Threaded then
        FThread.LeaveSection;
    end;
  end;
end;

function TCOMport.ReadLine : String;
const
  CBufferSize = 32;
var
  LBytesRead : LongWord;
  LCount     : LongWord;
  LPos       : LongWord;
begin
  Result := '';

  if IsOpen then
  begin
    if Threaded then
      FThread.EnterSection;
    try
      SetLength(Result, CBufferSize);

      LCount := 0;
      repeat
        Inc(LCount);
        ReadFile(FCOMhandle, Result[LCount], 1, LBytesRead, nil);

        LPos := Pos(FLineEnd, Result);
      until (LPos = Pred(LCount));

      LPos := Length(FLineEnd);
      SetLength(Result, (LCount - LPos));
    finally
      if Threaded then
        FThread.LeaveSection;
    end;
  end;
end;

procedure TCOMport.SetTimeouts(const AValue : TCommTimeouts);
begin
  FTimeouts := AValue;

  if IsOpen then
    SetCommTimeouts(FCOMhandle, FTimeouts);
end;

function TCOMport.PurgeBuffers(const ASpecification : LongInt = - 1) : Boolean;
begin
  Result := false;

  if IsOpen then
  begin
    if (ASpecification < 0) then
      Result := PurgeComm(FCOMhandle, PURGE_RXCLEAR or PURGE_TXCLEAR or PURGE_RXABORT or PURGE_TXABORT)
    else
      Result := PurgeComm(FCOMhandle, ASpecification);
  end;
end;

function TCOMport.IsOpen: Boolean;
begin
  Result := (FCOMhandle <> INVALID_HANDLE_VALUE);
end;

procedure TCOMport.CreateThread;
begin
  if (FThread = nil) then
  begin
    FThread := TCOMportThread.Create(true);
    try
      FThread.COMhandle  := FCOMhandle;
      FThread.OnReadChar := ReadChar;

      FThread.Resume;
    except
      DestroyThread;
    end;
  end;
end;

procedure TCOMport.DestroyThread;
begin
  if (FThread <> nil) then
  begin
    FThread.Terminate;
    FThread.Free;
    FThread := nil;
  end;
end;

{ TCOMportThread }

constructor TCOMportThread.Create(const ACreateSuspended : Boolean);
begin
  inherited Create(ACreateSuspended);

  FCOMhandle       := INVALID_HANDLE_VALUE;
  FCriticalSection := TCriticalSection.Create;
end;

destructor TCOMportThread.Destroy;
begin
  if (FCriticalSection <> nil) then
    FCriticalSection.Free;

  inherited;
end;

procedure TCOMportThread.EnterSection;
begin
  if (FCriticalSection <> nil) then
    FCriticalSection.Enter;
end;

procedure TCOMportThread.Execute;
var
  LChar  : Char;
  LRead  : LongWord;
begin
  while not(Terminated) do
  begin
    if (FCOMhandle <> INVALID_HANDLE_VALUE) then
    begin
      FCriticalSection.Enter;
      try
        ReadFile(FCOMhandle, LChar, 1, LRead, nil);
        if (LRead = 1) then
        begin
          if Assigned(FReadCharEvent) then
            FReadCharEvent(Self, LChar);
        end;
      finally
        FCriticalSection.Leave;
      end;
    end;
  end;
end;

procedure TCOMportThread.LeaveSection;
begin
  if (FCriticalSection <> nil) then
    FCriticalSection.Leave;
end;

end.
