unit SaveLog;
interface
uses
  Classes, SyncObjs;

type
  TLogSaveThread = class(TThread)
  private
    FTag: integer;
    FPort: string;
    FDataLine: string;
    FCriticalSection: TCriticalSection;
  protected
    procedure Execute; override;
  public
    constructor Create(const Port, DataLine: string; Tag: integer);
    destructor Destroy; override;
  end;

implementation

uses ComMainForm;

constructor TLogSaveThread.Create(const Port, DataLine: string; Tag: integer);
begin
  inherited Create(True);
  FTag := Tag;
  FPort := Port;
  FDataLine := DataLine;
  FCriticalSection := TCriticalSection.Create;
  Resume;
end;

 /////////////////////////////////////////////////////////////////////////////

destructor TLogSaveThread.Destroy;
begin
  FCriticalSection.Free;
  inherited Destroy;
end;

 /////////////////////////////////////////////////////////////////////////////

procedure TLogSaveThread.Execute;
var
  fFile: TextFile;
  Filename: string;
  FileStream: TFileStream;
begin
  Filename := LogDir + SessionDate + FPort + '.txt';
  try
    FileStream := TFileStream.Create(FileName, fmCreate);
    Mainform.GetComTerminal(FTag).SaveToStream(FileStream);
  finally
    FileStream.Free; // Free the stream
  end;
end;

end.

 