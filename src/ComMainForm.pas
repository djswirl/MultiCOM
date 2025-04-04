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
    sSkinManager1: TsSkinManager;
    sSkinProvider1: TsSkinProvider;
    StatusBar: TsStatusBar;
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
    FontSelect: TsComboBox;
    TimeSlider: TsSlider;
    ExspandBtn: TsSpeedButton;
    AutoLogBtn: TsSpeedButton;
    MenuImg: TImageList;
    BtnImg: TImageList;
    sSpeedButton2: TsSpeedButton;
    sSpeedButton3: TsSpeedButton;
    sSpeedButton4: TsSpeedButton;
    sSpeedButton5: TsSpeedButton;
    SaveTimer: TTimer;
    ShrinkBtn: TsSpeedButton;
    sSpeedButton7: TsSpeedButton;
    vFitBtn: TsSpeedButton;
    sSpeedButton6: TsSpeedButton;
    SOTSlider: TsSlider;
    ComPort0: TComPort;
    PortMemo0: TMemo;
    PortMemo1: TMemo;
    PortMemo2: TMemo;
    PortMemo3: TMemo;
    PortMemo4: TMemo;
    PortMemo5: TMemo;
    Timer1: TTimer;
    ConnectAllBtn: TsSpeedButton;
    sSpeedButton8: TsSpeedButton;
    sLabel1: TsLabel;
    sLabel2: TsLabel;
    sLabel3: TsLabel;
    sLabel4: TsLabel;
    sLabel5: TsLabel;
    sLabel6: TsLabel;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure OpenClose(Sender: TObject);
    function GetAppDir: string;
    procedure Settings(Sender: TObject);
    procedure SliderChange(Sender: TObject);
    procedure ResizeIconMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ResizeIconMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ResizeIconMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FontSelectCloseUp(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure ShrinkBtnClick(Sender: TObject);
    procedure PopExportClick(Sender: TObject);
    procedure PopClearClick(Sender: TObject);
    procedure ComPortException(Sender: TObject;
      TComException: TComExceptions; ComportMessage: string;
      WinError: Int64; WinMessage: string);
    procedure ErrorTimmerTimer(Sender: TObject);
    procedure AutoClick(Sender: TObject);
    procedure AutoLogBtnClick(Sender: TObject);
    procedure LogDirBtnClick(Sender: TObject);
    procedure ExspandBtnClick(Sender: TObject);
    procedure vFitBtnClick(Sender: TObject);
    procedure SOTSliderSliderChange(Sender: TObject);
    procedure SaveTimerTimer(Sender: TObject);
    procedure RxBuf(Sender: TObject; const Buffer; Count: Integer);
    procedure Timer1Timer(Sender: TObject);
    procedure ConnectAllBtnClick(Sender: TObject);

  private
    { Private declarations }
    FResizing: Boolean;
    FStartY: Integer;
    function GetLocalVersion: string;
    procedure UpDateStatusBar();
    procedure UpdateFont();
    procedure LoadConfig(Filename: string);
    procedure SaveConfig(Filename: string);


  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;
  VisAry: array[0..5] of bool;
  LogAry: array[0..5] of bool;
  PostStrAry: array[0..5] of string;
  LogRxStrAry: array[0..5] of string;
  LogStrAry: array[0..5] of TStringList;

  ComPort: array[0..5] of TComPort;
  ComSlider: array[0..5] of TsSlider;
  ComPanel: array[0..5] of TPanel;
  ComLabel: array[0..5] of TsLabel;
  ConnectBtn: array[0..5] of TSbutton;
  SettingsBtn: array[0..5] of TSbutton;
  AutoChk: array[0..5] of TsCheckbox;
  PortMemo: array[0..5] of TMemo;

  ConfigFile, LogDir, SessionDate: string;
  AutoLogOn: boolean;

implementation

uses Clipbrd, SaveLog;

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

procedure TMainForm.UpdateFont();
var
  x: integer;
begin
  for x := 0 to 5 do
    PortMemo[x].Font.Size := StrToInt(FontSelect.Items[FontSelect.ItemIndex]);
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

procedure TMainForm.LoadConfig(Filename: string);
var
  IniFile: TIniFile;
  x: integer;
begin
  IniFile := TIniFile.Create(Filename);
  try
    for x := 0 to 5 do
    begin
      ComPanel[x].Height := IniFile.ReadInteger('User', 'Com' + IntToStr(x) + 'Hght', 100);
      ComSlider[x].SliderOn := IniFile.ReadBool('User', 'Com' + IntToStr(x) + 'Enbl', true);
      VisAry[x] := ComSlider[x].SliderOn;

    end;

    TimeSlider.SliderOn := IniFile.ReadBool('User', 'Time', false);
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
      IniFile.WriteBool('User', 'Com' + IntToStr(x) + 'Enbl', ComSlider[x].SliderOn);
      IniFile.WriteInteger('User', 'Com' + IntToStr(x) + 'Hght', Companel[x].Height);
    end;

    IniFile.WriteBool('User', 'Time', TimeSlider.SliderOn);
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
var x: integer;
begin
  if not DirectoryExists(GetAppDir + '\Logs') then
    CreateDir(GetAppDir + '\Logs');

  for x := 0 to 5 do
  begin
    ComPort[x] := (FindComponent('ComPort' + IntToStr(x)) as TComPort);
    ComPort[x].OnRxBuf := RxBuf;
    ComSlider[x] := (FindComponent('ComSlider' + IntToStr(x)) as TsSlider);
    ComPanel[x] := (FindComponent('ComPanel' + IntToStr(x)) as TPanel);
    ComLabel[x] := (FindComponent('ComLabel' + IntToStr(x)) as TsLabel);
    ConnectBtn[x] := (FindComponent('ConnectBtn' + IntToStr(x)) as TSbutton);
    SettingsBtn[x] := (FindComponent('SettingsBtn' + IntToStr(x)) as TSbutton);
    AutoChk[x] := (FindComponent('AutoChk' + IntToStr(x)) as TsCheckbox);
    PortMemo[x] := (FindComponent('PortMemo' + IntToStr(x)) as TMemo);
    ComPanel[x].Align := alNone;
  end;




  for x := 5 to 0 do
  begin
    ComPanel[x].Align := alTop;
  end;

end;


/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.UpDateStatusBar();
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
  LogDir := GetAppDir + 'Logs\';
  LoadConfig(ConfigFile);
  UpDateStatusBar();
  UpdateFont();
  Statusbar.Panels[0].Text := 'V : ' + GetLocalVersion();


  for x := 0 to 5 do
  begin
    if FileExists(ConfigFile) then
      ComPort[x].LoadSettings(stIniFile, ConfigFile);
    ComPanel[x].SendToBack;
    ComPanel[x].Align := altop;
    LogStrAry[0] := TStringList.Create;
    ComLabel[x].Caption := ComPort[x].Port;
    StatusBar.Panels[x + 1].Text := ComPort[x].Port;
  end;
  SessionDate := FormatDateTime('yyyy-mm-dd_hh-nn', Now());
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
var x: integer;
begin
  for x := 0 to 5 do
    if (ComPort[x].Connected) then
      ComPort[x].close;
  SaveConfig(ConfigFile);
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.FormResize(Sender: TObject);
begin
  if (Height > Screen.DesktopHeight) then Height := Screen.DesktopHeight;
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
   // SaveTimer.Enabled := true; // enable the save timer
  end;
end;

  /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ConnectAllBtnClick(Sender: TObject);
var Slider: TsSlider;
var x: integer;
begin
  ConnectAllBtn.ImageIndex := 0; // reset image, only change again
  for x := 0 to 5 do
  begin
    if (ComSlider[x].SliderOn) then
    begin
      ConnectBtn[x].Click;
      if (ComPort[x].Connected) then
        ConnectAllBtn.ImageIndex := 1;
    end;
  end;
  timer1.Enabled := true;
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
    SaveTimer.Enabled := true; // enable the save timer
  end;
end;
  /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.LogDirBtnClick(Sender: TObject);
begin
  ShellExecute(Application.Handle, 'open', 'explorer.exe',
    PChar(GetAppDir + 'Logs'), nil, SW_NORMAL);
end;

////////////////////////////////////////////////////////////////////////////

procedure TMainForm.SliderChange(Sender: TObject);
var
  Slider: TsSlider;
  i, CurrentTop: Integer;
begin
  if Sender is TsSlider then
  begin
    Slider := Sender as TsSlider;
    ComPanel[Slider.Tag].Visible := ComSlider[Slider.Tag].SliderOn;
    VisAry[Slider.Tag] := Slider.SliderOn;

    // Ensure MenuPanel is always at the top
    MenuPanel.Top := 0;

    // Rearrange the other panels below the MenuPanel
    CurrentTop := MenuPanel.Top + MenuPanel.Height; // Start positioning below the menu
    for i := 0 to High(ComPanel) do
    begin
      if ComPanel[i].Visible then
      begin
        ComPanel[i].Top := CurrentTop;
        Inc(CurrentTop, ComPanel[i].Height); // Move down for the next panel
      end;
    end;
  end;
end;


/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.OpenClose(Sender: TObject);
var btn: TsButton;
begin
  application.ProcessMessages;
  if Sender is TsButton then
  begin
    Timer1.Enabled := true;
    Btn := Sender as TsButton;
    if (ComPort[Btn.Tag] <> nil) then
    begin
      PostStrAry[Btn.Tag] := ComPort[Btn.Tag].Port;
      if ComPort[Btn.Tag].Connected then
      begin
        ComPort[Btn.Tag].Close;
        ComLabel[Btn.Tag].Font.Color := clRed;
        ConnectBtn[Btn.Tag].Caption := 'Connect';
        ConnectBtn[Btn.Tag].ImageIndex := 0;
        ConnectBtn[Btn.Tag].Down := False;
      end else begin
        try
          ComPort[Btn.Tag].open;
          ComLabel[Btn.Tag].Font.Color := $0033CC00;
          ConnectBtn[Btn.Tag].Caption := 'Disconnect';
          ConnectBtn[Btn.Tag].ImageIndex := 1;
          ConnectBtn[Btn.Tag].Down := True;
          LogRxStrAry[Tag] := '';
        except
          on E: Exception do
          begin
            StatusBar.Panels[Tag + 1].Text := 'Error: + ' + E.Message;
          end;
        end;
      end;
      application.ProcessMessages;
    end;
  end;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TMainForm.ComPortException(Sender: TObject;
  TComException: TComExceptions; ComportMessage: string; WinError: Int64;
  WinMessage: string);
var
  CPort: TComPort;
begin
  // Check if the Sender is a TComPort
  if Sender is TComPort then
  begin
    CPort := TComPort(Sender); // Cast Sender to TComPort

    // Use the Tag property to identify the triggering ComPort
  //  if Assigned(ConnectBtn[ComPort.Tag]) then
  //  begin
//ComPort[CPort.Tag].Close;
   // ConnectBtn[CPort.Tag].Click;

    ComLabel[CPort.Tag].Font.Color := clRed;
    ConnectBtn[CPort.Tag].Caption := 'Connect';
    ConnectBtn[CPort.Tag].ImageIndex := 0;
    ConnectBtn[CPort.Tag].Down := False;
 //   showmessage('panic');
    application.ProcessMessages;
 //   end;

    // Update the StatusBar with the error information
    Statusbar.Panels[CPort.Tag + 1].Text := Format('%s: %s', [ComPort[CPort.Tag].Port, ComportMessage]);
    Statusbar.Panels[7].Text := Format('ERROR on COM%d: %s (%d)', [CPort.Tag, ComportMessage, WinError]);
  end
  else
  begin
    // Handle unexpected Sender types
    Statusbar.Panels[7].Text := 'ERROR: Unknown COM port exception';
  end;

  // Enable the error timer to clear the StatusBar after a delay
  ErrorTimmer.Enabled := True;
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

    ComPort[SenderBtn.Tag].ShowSetupDialog;
    PortMemo[SenderBtn.Tag].clear; ;

    ComPort[SenderBtn.Tag].StoreSettings(stIniFile, GetAppDir + 'config.ini');
    ComLabel[SenderBtn.Tag].Caption := ComPort[SenderBtn.Tag].Port;
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
    Companel[Img.Tag].Height := Companel[Img.Tag].Height + DeltaY;
    FStartY := Mouse.CursorPos.Y;
  end;
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
      Companel[x].Height := PHeight div DivCount;
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
  x, HeadFoot: integer;
  MHeight: integer;
begin
  if (Height > Screen.WorkAreaHeight) then Height := Screen.WorkAreaHeight;
  if (Width > Screen.WorkAreaWidth) then Width := Screen.WorkAreaWidth;
  MHeight := 0;
  for x := 0 to 5 do
  begin
    if (VisAry[x] = true) then
    begin
      MHeight := MHeight + Companel[x].Height;
    end;
  end;
  HeadFoot := MenuPanel.Height + StatusBar.Height;
  MainForm.ClientHeight := MHeight + HeadFoot;

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
  MHeight: integer;
begin
  if (Height > Screen.WorkAreaHeight) then Height := Screen.WorkAreaHeight;
  if (Width > Screen.WorkAreaWidth) then Width := Screen.WorkAreaWidth;
  MHeight := 0;
  for x := 0 to 5 do
  begin
    if (VisAry[x] = true) then
      MHeight := MHeight + Companel[x].Height;

  end;
  HeadFoot := MenuPanel.Height + StatusBar.Height;
  MainForm.ClientHeight := MHeight + HeadFoot;

  if (Height < Screen.WorkAreaHeight) then
  begin
    AutoSize := true;
    AutoSize := false;
  end;
  update;
end;

 /////////////////////////////////////////////////////////////////////////////


procedure TMainForm.SaveTimerTimer(Sender: TObject);
var x: integer;
  Filename: string;
  FileStream: TFileStream;
begin
  application.ProcessMessages;
  for x := 0 to 5 do
  begin
    if ((ComPort[x].Connected) and (LogAry[x])) then
    begin
      FileName := GetAppDir + 'logs\' + SessionDate + '_' + ComPort[x].Port + '_Data.txt';

      // Create a TFileStream to write the contents of the terminal to a file
      FileStream := TFileStream.Create(FileName, fmCreate);
      try
        // Save the contents of the memo to the stream
        PortMemo[x].Lines.SaveToStream(FileStream);
      finally
        FileStream.Free; // Free the stream
      end;
    end;
  end;
end;

 /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.RxBuf(Sender: TObject; const Buffer; Count: Integer);
var
  Data: AnsiString;
  ReceivedData: array of Byte;
  ComSender: tComPort;
  Tag, i: Integer;
  TerminatorPos: Integer;
  RawOutput: string;
begin
  ComSender := Sender as tComPort;
  Tag := ComSender.Tag;

  // Check if Count is valid
  if Count <= 0 then Exit;

  StatusBar.Panels[Tag + 1].Text := IntToStr(Count) + ' Bytes received';

  // Allocate memory and copy data
  SetLength(ReceivedData, Count);
  Move(Buffer, ReceivedData[0], Count);


  // Convert raw data to a string and append it to the buffer
  SetString(Data, PAnsiChar(@Buffer), Count);
  LogRxStrAry[Tag] := LogRxStrAry[Tag] + Data;

  // Check for complete messages terminated by #10 (LF) or #13#10 (CRLF)
  repeat
    TerminatorPos := Pos(#10, LogRxStrAry[Tag]); // Adjust to match your terminator
    if TerminatorPos > 0 then
    begin
      // Extract the complete message up to the terminator
      if (TimeSlider.SliderOn) then
        PortMemo[tag].Lines.Append(FormatDateTime('hh:nn:ss.zzz', Now()) + ' - ' + Copy(LogRxStrAry[Tag], 1, TerminatorPos - 1))
      else

        PortMemo[tag].Lines.Append(Copy(LogRxStrAry[Tag], 1, TerminatorPos - 1));

      // Remove the processed message from the buffer
      Delete(LogRxStrAry[Tag], 1, TerminatorPos);
    end;
  until TerminatorPos = 0;
end;

 /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.PopExportClick(Sender: TObject);
var
  Memo: TMemo;
  FileStream: TFileStream;
begin
  // Check if the PopupMenu's PopupComponent is a TComTerminal
  if ExportPopup.PopupComponent is TMemo then
  begin
    // Access the TComTerminal and get its Tag
    Memo := TMemo(ExportPopup.PopupComponent);
    SaveDialog.FileName := ComPort[Tag].Port + '-Data.txt';

    if SaveDialog.Execute then
    begin
      // Create a TFileStream to write the contents of the terminal to a file
      FileStream := TFileStream.Create(SaveDialog.FileName, fmCreate);
      try
        // Save the contents of the terminal to the stream
        PortMemo[tag].Lines.SaveToStream(FileStream);
      finally
        FileStream.Free; // Free the stream
      end;
    end;
  end;
end;
 /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.PopClearClick(Sender: TObject);
var
  Memo: TMemo;
begin
  // Check if the PopupMenu's PopupComponent is a TMemo
  if ExportPopup.PopupComponent is TMemo then
  begin
    // Access the TComTerminal and get its Tag
    Memo := TMemo(ExportPopup.PopupComponent);
    Memo.Clear;
  end

end;

  /////////////////////////////////////////////////////////////////////////////

procedure TMainForm.SOTSliderSliderChange(Sender: TObject);
begin
  if (SOTSlider.SliderOn) then
    SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NoMove or SWP_NoSize)
  else
    SetWindowPos(Handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NoMove or SWP_NoSize);
end;

 /////////////////////////////////////////////////////////////////////////////



procedure TMainForm.Timer1Timer(Sender: TObject);
var x: integer;
begin
   exit;
  Timer1.Enabled := false; ConnectAllBtn.ImageIndex := 0;
  for x := 0 to 5 do
  begin
    if (ComSlider[x].SliderOn) then
    begin
      ComPort[x].Close;
      ComLabel[x].Font.Color := clRed;
      ConnectBtn[x].Caption := 'Connect';
      ConnectBtn[x].ImageIndex := 0;
      ConnectBtn[x].Down := False;
    end;
  end;

end;
end.

