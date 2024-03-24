program MultiCom;

uses
  Forms,
  ComMainForm in 'ComMainForm.pas' {MainForm},
  SaveLog in 'SaveLog.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'MultiCOM';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.

