unit ComMainForm;


// change in CPortCtl
// FFontWidth := Metrics.tmAveCharWidth.tmMaxCharWidth;
// to
// FFontWidth := Metrics.tmAveCharWidth;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, CPort, CPortCtl, sSkinProvider, sButton, sPanel,
  sSkinManager, sMemo, sSplitter, sLabel, acSlider, acPNG, acImage,
  sComboBox, IniFiles, DateUtils, ComCtrls, sStatusBar, Menus, sDialogs,
  sCheckBox, ImgList, ToolWin, acCoolBar, sToolBar, acHeaderControl,
  Buttons, sSpeedButton, sBitBtn, ShellAPI, acArcControls;

type
  TMainForm = class(TForm)
    ComPort1: TComPort;
    ComPort2: TComPort;
    ComPort3: TComPort;
    ComPort4: TComPort;
    ComPort5: TComPort;
    MenuPanel: TsPanel;
    ComPanel0: TPanel;
    Panel2: TPanel;
    sImage7: TsImage;
    ComLabel0: TsLabel;
    ConnectBtn0: TsButton;
    SettingsBtn0: TsButton;
    ComPanel5: TPanel;
    sImage8: TsImage;
    Panel3: TPanel;
    ComLabel5: TsLabel;
    ConnectBtn5: TsButton;
    SettingsBtn5: TsButton;
    ComPanel4: TPanel;
    sImage9: TsImage;
    Panel5: TPanel;
    ComLabel4: TsLabel;
    ConnectBtn4: TsButton;
    SettingsBtn4: TsButton;
    ComPanel3: TPanel;
    sImage10: TsImage;
    Panel7: TPanel;
    ComLabel3: TsLabel;
    ConnectBtn3: TsButton;
    SettingsBtn3: TsButton;
    ComPanel2: TPanel;
    sImage11: TsImage;
    Panel9: TPanel;
    ComLabel2: TsLabel;
    ConnectBtn2: TsButton;
    SettingsBtn2: TsButton;
    ComPanel1: TPanel;
    sImage12: TsImage;
    Panel11: TPanel;
    ComLabel1: TsLabel;
    ConnectBtn1: TsButton;
    SettingsBtn1: TsButton;
    ComTerminal1: TComTerminal;
    ComTerminal2: TComTerminal;
    ComTerminal3: TComTerminal;
    ComTerminal4: TComTerminal;
    ComTerminal5: TComTerminal;
    sSkinManager1: TsSkinManager;
    sSkinProvider1: TsSkinProvider;
    StatusBar: TsStatusBar;
    ComTerminal0: TComTerminal;
    ComPort0: TComPort;
    ExportPopup: TPopupMenu;
    PopExport: TMenuItem;
    SaveDialog: TsSaveDialog;
    AutoChk0: TsCheckBox;
    AutoChk1: TsCheckBox;
    AutoChk2: TsCheckBox;
    AutoChk3: TsCheckBox;
    AutoChk4: TsCheckBox;
    AutoChk5: TsCheckBox;
    PopClear: TMenuItem;
    ErrorTimmer: TTimer;
    sPanel1: TsPanel;
    ComSlider0: TsSlider;
    ComSlider1: TsSlider;
    ComSlider2: TsSlider;
    ComSlider3: TsSlider;
    ComSlider4: TsSlider;
    ComSlider5: TsSlider;
    LogDirBtn: TsSpeedButton;
    sPanel2: TsPanel;
    LFSilder: TsSlider;
    emulationCombo: TsComboBox;
    FontSelect: TsComboBox;
    TimeSlider: TsSlider;
    TABSlider: TsSlider;
    ExspandBtn: TsSpeedButton;
    AutoLogBtn: TsSpeedButton;
    MenuImg: TImageList;
    BtnImg: TImageList;
    sSpeedButton2: TsSpeedButton;
    sSpeedButton3: TsSpeedButton;
    sSpeedButton4: TsSpeedButton;
    sSpeedButton5: TsSpeedButton;
    Timer0: TTimer;
    ShrinkBtn: TsSpeedButton;
    sSpeedButton7: TsSpeedButton;
    vFitBtn: TsSpeedButton;
    sSpeedButton6: TsSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure OpenClose(Sender: TObject);
    procedure Settings(Sender: TObject);
    procedure SliderChange(Sender: TObject);
    procedure ResizeIconMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ResizeIconMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ResizeIconMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure LFSilderSliderChange(Sender: TObject);
    procedure TABSliderSliderChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ComTerminalResize(Sender: TObject);
    procedure FontSelectCloseUp(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure ShrinkBtnClick(Sender: TObject);
    procedure ComPortRxChar(Sender: TObject; Count: Integer);
    procedure PopExportClick(Sender: TObject);
    procedure SaveToLog(SessionDate: string; Port: string; DataLine: string);
    procedure PopClearClick(Sender: TObject);
    procedure emulationComboCloseUp(Sender: TObject);
    procedure ComPortException(Sender: TObject;
      TComException: TComExceptions; ComportMessage: string;
      WinError: Int64; WinMessage: string);
    procedure ErrorTimmerTimer(Sender: TObject);
    function RemoveEscapeCodes(const InputStr: string): string;
    procedure AutoClick(Sender: TObject);
    procedure AutoLogBtnClick(Sender: TObject);
    procedure LogDirBtnClick(Sender: TObject);
    procedure ExspandBtnClick(Sender: TObject);
    procedure vFitBtnClick(Sender: TObject);
  private
    { Private declarations }
    FResizing: Boolean;
    FStartY: Integer;
    function GetLocalVersion: string;
    function GetAppDir: string;
    function GetConnectBtn(idx: integer): TsButton;
    function GetComPort(idx: integer): TComPort;
    function GetComTerminal(idx: integer): TComTerminal;
    function GetComSlider(idx: integer): TsSlider;
    function GetCompanel(idx: integer): tPanel;
    function GetComLabel(idx: integer): TsLabel;
    procedure UpDateStatusBar();
    procedure UpdateFont();
    procedure UpdateEmulation();
    procedure LoadConfig(Filename: string);
    procedure SaveConfig(Filename: string);


  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;
  VisAry: array[0..5] of bool;
  LogAry: array[0..5] of bool;
  RxAry: array[0..5] of string;
  ConfigFile, SessionDate: string;
  AutoLogOn: boolean;
implementation

uses Clipbrd;

{$R *.DFM}



function TMainForm.GetLocalVersion: string;
var
  VerInfoSize: DWORD;
  VerInfo: Pointer;
  VerValueSize: DWORD;
  VerValue: PVSFixedFileInfo;
  Dummy: DWORD;
begin
  VerInfoSize := GetFileVersionInfoSize(PChar(ParamStr(0)), Dummy);
  GetMem(VerInfo, VerInfoSize);
  GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo);
  VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
  with VerValue^ do
  begin
    Result := IntToStr(dwFileVersionMS shr 16);
    Result := Result + '.' + IntToStr(dwFileVersionMS and $FFFF);
    Result := Result + '.' + IntToStr(dwFileVersionLS shr 16);
    Result := Result + '.' + IntToStr(dwFileVersionLS and $FFFF);
  end;
  FreeMem(VerInfo, VerInfoSize);
end;

 /////////////////////////////////////////////////////////////////////////////

function TMainForm.GetAppDir: string;
begin
  Result := ExtractFilePath(Paramstr(0));
  if Result[Length(result)] <> '\' then
    Result := Result + '\';
end;

/////////////////////////////////////////////////////////////////////////////

function TMainForm.GetConnectBtn(idx: integer): TsButton;
var Btn: TsButton;
begin
  Btn := FindComponent('ConnectBtn' + IntToStr(idx)) as TsButton;
  if (Btn <> nil) then Result := Btn else Result := nil;
end;

/////////////////////////////////////////////////////////////////////////////

function TMainForm.GetComPort(idx: integer): TComPort;
var ComPort: TComPort;
begin
  ComPort := FindComponent('ComPort' + IntToStr(idx)) as TComPort;
  if (ComPort <> nil) then Result := ComPort else Result := nil;
end;

/////////////////////////////////////////////////////////////////////////////

function TMainForm.GetComTerminal(idx: integer): TComTerminal;
var ComTerminal: TComTerminal;
begin
  ComTerminal := FindComponent('ComTerminal' + IntToStr(idx)) as TComTerminal;
  if (ComTerminal <> nil) then Result := ComTerminal else Result := nil;
end;
/////////////////////////////////////////////////////////////////////////////

function TMainForm.GetComSlider(idx: integer): TsSlider;
var Slider: TsSlider;
begin
  Slider := (FindComponent('ComSlider' + IntToStr(idx)) as TsSlider);
  if (Slider <> nil) then Result := Slider else Result := nil;
end;
/////////////////////////////////////////////////////////////////////////////

function TMainForm.GetCompanel(idx: integer): tPanel;
var Panel: tPanel;
begin
  Panel := (FindComponent('ComPanel' + IntToStr(idx)) as tPanel);
  if (Panel <> nil) then Result := Panel else Result := nil;
end;
/////////////////////////////////////////////////////////////////////////////

function TMainForm.GetComLabel(idx: integer): TsLabel;
var sLabel: TsLabel;
begin
  sLabel := (FindComponent('ComLabel' + IntToStr(idx)) as TsLabel);
  if (sLabel <> nil) then Result := sLabel else Result := nil;
end;
/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.UpdateFont();
var
  x: integer;
begin
  for x := 0 to 5 do
    GetComTerminal(x).Font.Size := StrToInt(FontSelect.Items[FontSelect.ItemIndex]);
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.UpdateEmulation();
var x: integer;
  CustomComTerminal: TCustomComTerminal;

begin
  for x := 0 to 5 do
  begin
    case emulationCombo.ItemIndex of
      0: GetComTerminal(x).Emulation := teNone;
      1: GetComTerminal(x).Emulation := teVT100orANSI;
      2: GetComTerminal(x).Emulation := teVT52;
    end;
  end;
 /// CustomComTerminal := TCustomComTerminal.Create(self);
 // CustomComTerminal.EscapeCodes
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ResizeIconMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  FResizing := True;
  FStartY := Mouse.CursorPos.Y;
end;
/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ResizeIconMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  FResizing := False;
end;
/////////////////////////////////////////////////////////////////////////////

function PosEx(const SubStr, S: string; Offset: Integer = 1): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := Offset to Length(S) - Length(SubStr) + 1 do
    if Copy(S, i, Length(SubStr)) = SubStr then
    begin
      Result := i;
      Break;
    end;
end;
/////////////////////////////////////////////////////////////////////////////

function TMainForm.RemoveEscapeCodes(const InputStr: string): string;
var
  EscapeCodeStart, EscapeCodeEnd: Integer;
begin
  Result := InputStr;
  EscapeCodeStart := Pos(#27, Result);
  while EscapeCodeStart > 0 do
  begin
    EscapeCodeEnd := PosEx('m', Result, EscapeCodeStart);
    if EscapeCodeEnd > 0 then
      Delete(Result, EscapeCodeStart, EscapeCodeEnd - EscapeCodeStart + 1);
    EscapeCodeStart := PosEx(#27, Result, EscapeCodeStart);
  end;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.SaveToLog(SessionDate: string; Port: string; DataLine: string);
var
  fFile: TextFile;
  Filename: string;
begin
  DataLine := StringReplace(DataLine, #10, '', [rfReplaceAll, rfIgnoreCase]);
  DataLine := StringReplace(DataLine, #13, '', [rfReplaceAll, rfIgnoreCase]);
  DataLine := RemoveEscapeCodes(DataLine);

  Filename := GetAppDir + 'Logs\' + SessionDate + Port + '.txt';
  AssignFile(fFile, Filename);
  if FileExists(Filename) <> true then
    Rewrite(fFile)
  else
    Append(fFile);
  WriteLn(fFile, DataLine);
  CloseFile(fFile)
end;


/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.LoadConfig(Filename: string);
var
  IniFile: TIniFile;
  x: integer;
begin
  IniFile := TIniFile.Create(Filename);
  try
    for x := 0 to 5 do
    begin
      GetComSlider(x).SliderOn := IniFile.ReadBool('User', 'Com' + IntToStr(x) + 'Enbl', true);
      VisAry[x] := GetComSlider(x).SliderOn;
      GetCompanel(x).Height := IniFile.ReadInteger('User', 'Com' + IntToStr(x) + 'Hght', 100);
    end;

    TABSlider.SliderOn := IniFile.ReadBool('User', 'Tab', true);
    LFSilder.SliderOn := IniFile.ReadBool('User', 'LF', false);
    TimeSlider.SliderOn := IniFile.ReadBool('User', 'Time', false);
    emulationCombo.ItemIndex := IniFile.ReadInteger('User', 'Emu', 0);
    FontSelect.ItemIndex := IniFile.ReadInteger('User', 'Font', 0);

    Left := IniFile.ReadInteger('User', 'L', Left);
    Top := IniFile.ReadInteger('User', 'T', Top);
    Width := IniFile.ReadInteger('User', 'W', Width);
    Height := IniFile.ReadInteger('User', 'H', Height)

  finally
    IniFile.Free;
  end;
end;
/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.SaveConfig(Filename: string);
var
  IniFile: TIniFile;
  x: integer;
begin
  IniFile := TIniFile.Create(Filename);
  try
    for x := 0 to 5 do
    begin
      IniFile.WriteBool('User', 'Com' + IntToStr(x) + 'Enbl', GetComSlider(x).SliderOn);
      IniFile.WriteInteger('User', 'Com' + IntToStr(x) + 'Hght', GetCompanel(x).Height);
    end;

    IniFile.WriteBool('User', 'Tab', TABSlider.SliderOn);
    IniFile.WriteBool('User', 'LF', LFSilder.SliderOn);
    IniFile.WriteBool('User', 'Time', TimeSlider.SliderOn);

    IniFile.WriteInteger('User', 'Emu', emulationCombo.ItemIndex);
    IniFile.WriteInteger('User', 'Font', FontSelect.ItemIndex);
    IniFile.WriteInteger('User', 'L', Left);
    IniFile.WriteInteger('User', 'T', Top);
    IniFile.WriteInteger('User', 'W', Width);
    IniFile.WriteInteger('User', 'H', Height);

  finally
    IniFile.Free;
  end;
end;
/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.FormCreate(Sender: TObject);
begin
  if not DirectoryExists(GetAppDir + '\Logs') then
    CreateDir(GetAppDir + '\Logs')
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.UpDateStatusBar();
var x: integer;
begin
  StatusBar.Panels[7].Width := Width - 880;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.FormShow(Sender: TObject);
var
  x: integer;
begin
  AutoLogOn := false;
  ConfigFile := GetAppDir + 'config.ini';
  LoadConfig(ConfigFile);
  UpDateStatusBar();
  UpdateEmulation();
  UpdateFont();
  Statusbar.Panels[0].Text := 'V : ' + GetLocalVersion();

  for x := 0 to 5 do
  begin
    if FileExists(ConfigFile) then
      GetComPort(x).LoadSettings(stIniFile, ConfigFile);
    GetComLabel(x).Caption := GetComPort(x).Port;
    StatusBar.Panels[x + 1].Text := GetComPort(x).Port;
  end;
  SessionDate := FormatDateTime('yyyy-mm-dd-hh-nn-', Now());

end;
/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
var x: integer;
begin
  for x := 0 to 5 do
    if (GetCOmPort(x).Connected) then
      GetComPort(x).close;
  SaveConfig(ConfigFile);
end;
/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.FormResize(Sender: TObject);
begin
  if (Height > Screen.DesktopHeight) then Height := Screen.DesktopHeight;
  ComTerminalResize(nil);
  UpDateStatusBar();
end;
   /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.AutoClick(Sender: TObject);
var ChkSender, CheckBox: TsCheckBox;
begin
  if Sender is TsCheckBox then
  begin
    ChkSender := Sender as TsCheckBox;
    CheckBox := FindComponent('AutoChk' + IntToStr(ChkSender.Tag)) as TsCheckBox;
    LogAry[ChkSender.Tag] := CheckBox.Checked;
  end;
end;


  /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.AutoLogBtnClick(Sender: TObject);
var CheckBox: TsCheckBox;
var x: integer;
begin
  AutoLogOn := not AutoLogOn;
  for x := 0 to 5 do
  begin
    CheckBox := FindComponent('AutoChk' + IntToStr(x)) as TsCheckBox;
    CheckBox.Checked := AutoLogOn;
  end;
end;
  /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.LogDirBtnClick(Sender: TObject);
begin
  ShellExecute(Application.Handle, 'open', 'explorer.exe',
    PChar(GetAppDir + 'Logs'), nil, SW_NORMAL);
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.LFSilderSliderChange(Sender: TObject);
var
  x: integer;
begin
  for x := 0 to 5 do
    GetComTerminal(x).AppendLF := LFSilder.SliderOn
end;

////////////////////////////////////////////////////////////////////////////

procedure TMainForm.SliderChange(Sender: TObject);
var
  Slider: TsSlider;
begin
  if Sender is TsSlider then
  begin
    Slider := Sender as TsSlider;
    GetCompanel(Slider.Tag).Visible := Slider.SliderOn;
    VisAry[Slider.Tag] := Slider.SliderOn;
  end;
end;
/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.TABSliderSliderChange(Sender: TObject);
var
  x: integer;
begin
  for x := 0 to 5 do
    GetComTerminal(x).WantTab := TABSlider.SliderOn;
end;
/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.OpenClose(Sender: TObject);
var btn: TsButton;
begin
  if Sender is TsButton then
  begin
    Btn := Sender as TsButton;
    if (GetComPort(Btn.Tag) <> nil) then
    begin
      if GetComPort(Btn.Tag).Connected then
      begin
        GetComPort(Btn.Tag).Close;
        GetComLabel(Btn.Tag).Font.Color := clRed;
        Btn.Caption := 'Connect';
        GetConnectBtn(Btn.Tag).ImageIndex := 0;
        GetConnectBtn(Btn.Tag).Down := False;
        Timer0.Enabled := false;
      end else begin
        GetComPort(Btn.Tag).Open;
        GetComLabel(Btn.Tag).Font.Color := $0033CC00;
        Btn.Caption := 'Disconnect';
        GetConnectBtn(Btn.Tag).ImageIndex := 1;
        GetConnectBtn(Btn.Tag).Down := True;
        Timer0.Enabled := True;
      end;
    end;
  //  AutoBtn.SetFocus;
  end;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ComPortException(Sender: TObject;
  TComException: TComExceptions; ComportMessage: string; WinError: Int64;
  WinMessage: string);
begin
  case WinError of
    2: begin
        Dialogs.MessageDlg(ComportMessage, mtWarning, [mbOK], 0);

      end;
  else
    Statusbar.Panels[7].Text := 'ERROR : ' + ComportMessage + ' (' + IntToStr(WinError) + ')'; ErrorTimmer.Enabled := True;
  end;
 // abort;

end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ErrorTimmerTimer(Sender: TObject);
begin
  Statusbar.Panels[7].Text := '';
  ErrorTimmer.Enabled := False;
end;
/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.Settings(Sender: TObject);
var
  SenderBtn: TsButton;
begin
  if Sender is TButton then
  begin
    SenderBtn := Sender as TsButton;
    GetComPort(SenderBtn.Tag).ShowSetupDialog;
    GetComPort(SenderBtn.Tag).StoreSettings(stIniFile, GetAppDir + 'config.ini');
    GetComLabel(SenderBtn.Tag).Caption := GetComPort(SenderBtn.Tag).Port;
  end;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ResizeIconMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
var
  DeltaY: Integer;
  Img: TsImage;
begin
  if FResizing then
  begin
    DeltaY := Mouse.CursorPos.Y - FStartY;
    Img := Sender as TsImage;
    GetCompanel(Img.Tag).Height := GetCompanel(Img.Tag).Height + DeltaY;
    FStartY := Mouse.CursorPos.Y;
  end;
end;

 /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ComTerminalResize(Sender: TObject);
var
  x: integer;
begin
 // for x := 0 to 5 do
 //   GetComTerminal(x).Rows := (GetCompanel(x).Height div GetComTerminal(x).Font.Size) - 6;
end;
 /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.emulationComboCloseUp(Sender: TObject);
begin
  UpdateEmulation();
end;
 /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.FontSelectCloseUp(Sender: TObject);
begin
  UpdateFont();
end;

/////////////////////////////////////////////////////////////////////////////
 procedure TMainForm.vFitBtnClick(Sender: TObject);
var
  x, HeadFoot, PHeight: integer;
  DivCount: integer;
begin
  if (Height > Screen.WorkAreaHeight) then Height := Screen.WorkAreaHeight;
  if (Width > Screen.WorkAreaWidth) then Width := Screen.WorkAreaWidth;
  DivCount := 0;
  for x := 0 to 5 do
  begin
    if (VisAry[x] = true) then DivCount := DivCount + 1;
  end;
  HeadFoot := MenuPanel.Height + StatusBar.Height;
  PHeight := MainForm.ClientHeight - HeadFoot;
  if (DivCount > 0) then
  begin
    for x := 0 to 5 do
      GetCompanel(x).Height := PHeight div DivCount;
  end;

  if (Height < Screen.WorkAreaHeight) then
  begin
    AutoSize := true;
    AutoSize := false;
  end;
  update;

end;

/////////////////////////////////////////////////////////////////////////////
procedure TMainForm.ExspandBtnClick(Sender: TObject);
var
  x, HeadFoot, PHeight: integer;
  DivCount, MHeight: integer;
begin
 if (Height > Screen.WorkAreaHeight) then Height := Screen.WorkAreaHeight;
 if (Width > Screen.WorkAreaWidth) then Width := Screen.WorkAreaWidth;
  DivCount := 0;
  MHeight :=0;
  for x := 0 to 5 do
  begin
    if (VisAry[x] = true) then
    begin
    MHeight := MHeight + GetCompanel(x).Height;
    DivCount := DivCount + 1;
    end;

  end;
  HeadFoot := MenuPanel.Height + StatusBar.Height;
  MainForm.ClientHeight:= MHeight + HeadFoot;

  if (Height > Screen.WorkAreaHeight) then
  begin
vFitBtn.Click;
  end;


  update;
end;

  /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ShrinkBtnClick(Sender: TObject);
var
  x, HeadFoot, PHeight: integer;
  DivCount, MHeight: integer;
begin
 if (Height > Screen.WorkAreaHeight) then Height := Screen.WorkAreaHeight;
 if (Width > Screen.WorkAreaWidth) then Width := Screen.WorkAreaWidth;
  DivCount := 0;
  MHeight :=0;
  for x := 0 to 5 do
  begin
    if (VisAry[x] = true) then
    MHeight := MHeight + GetCompanel(x).Height;

  end;
  HeadFoot := MenuPanel.Height + StatusBar.Height;
  MainForm.ClientHeight:= MHeight + HeadFoot;

  if (Height < Screen.WorkAreaHeight) then
  begin
 AutoSize := true;
AutoSize := false;
  end;


  update;

end;

 /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ComPortRxChar(Sender: TObject; Count: Integer);
var
  Str: string;
  ComSender: tComPort;
begin
  ComSender := Sender as tComPort;
  GetComPort(ComSender.Tag).ReadStr(Str, Count);

  RxAry[ComSender.Tag] := RxAry[ComSender.Tag] + Str;

  if ((Pos(#0, RxAry[ComSender.Tag]) > 0) or (Pos(#13 + #10, RxAry[ComSender.Tag]) > 0)) then
  begin
    StatusBar.Panels[ComSender.Tag + 1].Text := GetComPort(ComSender.Tag).Port + ': ' + IntToStr(Length(RxAry[ComSender.Tag])) + ' Bytes recv';
    if (TimeSlider.SliderOn) then
    begin
      RxAry[ComSender.Tag] := FormatDateTime('hh:nn:ss.zzz', Now()) + ' : ' + RxAry[ComSender.Tag];
    end;
    GetComTerminal(ComSender.Tag).WriteStr(RxAry[ComSender.Tag]);

    if (LogAry[ComSender.Tag]) then SaveToLog(SessionDate, GetComPort(ComSender.Tag).Port, RxAry[ComSender.Tag]);
    RxAry[ComSender.Tag] := '';
  end;
end;
 /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.PopExportClick(Sender: TObject);
var
  Terminal: TComTerminal;
  FileStream: TFileStream;
begin
  // Check if the PopupMenu's PopupComponent is a TComTerminal
  if ExportPopup.PopupComponent is TComTerminal then
  begin
    // Access the TComTerminal and get its Tag
    Terminal := TComTerminal(ExportPopup.PopupComponent);
    SaveDialog.FileName := GetComPort(Terminal.Tag).Port + '-Data.txt';

    if SaveDialog.Execute then
    begin
      // Create a TFileStream to write the contents of the terminal to a file
      FileStream := TFileStream.Create(SaveDialog.FileName, fmCreate);
      try
        // Save the contents of the terminal to the stream
        Terminal.SaveToStream(FileStream);
      finally
        FileStream.Free; // Free the stream
      end;
    end;
  end;
end;
 /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.PopClearClick(Sender: TObject);
var
  Terminal: TComTerminal;
begin
  // Check if the PopupMenu's PopupComponent is a TComTerminal
  if ExportPopup.PopupComponent is TComTerminal then
  begin
    // Access the TComTerminal and get its Tag
    Terminal := TComTerminal(ExportPopup.PopupComponent);
    Terminal.ClearScreen;
  end

end;





end.

