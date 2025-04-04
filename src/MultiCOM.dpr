program MultiCom;

uses
  Forms,
  ComMainForm in 'ComMainForm.pas' {MainForm};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'MultiCOM';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.

