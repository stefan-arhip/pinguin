unit uMain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ComCtrls,
  StdCtrls, Spin, ComObj, Variants, ActiveX, Windows, CommCtrl;

type

  { TfMain }

  TfMain = class(TForm)
    ImageList1: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    ListView1: TListView;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    seBuffer: TSpinEdit;
    seTimeout: TSpinEdit;
    seRetries: TSpinEdit;
    seMaxThreads: TSpinEdit;
    StatusBar1: TStatusBar;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    ToolBar1: TToolBar;
    ToolBar2: TToolBar;
    ToolButton2: TToolButton;
    ToolButton1: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    procedure FormCreate(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

  TPingThread = class(TThread)
    Id, Buffer, Timeout, Retries: integer;
    Computer: string;
    procedure Execute; override;
  end;

var
  fMain: TfMain;

implementation

{$R *.lfm}

{ TfMain }

var
  ActiveThreads: integer;
  StartTime: double;

function WMIPing(const Address: string; const BufferSize, Timeout: word): integer;
const
  WbemUser = '';
  WbemPassword = '';
  WbemComputer = 'localhost';
  wbemFlagForwardOnly = $00000020;
var
  FSWbemLocator: olevariant;
  FWMIService: olevariant;
  FWbemObjectSet: olevariant;
  FWbemObject: olevariant;
  CeltFetched: ULONG;
  oEnum: ActiveX.IEnumvariant;
  FWbemQuery: string[250];
begin
  Result := -1;
  CoInitialize(nil);
  try
    FSWbemLocator := ComObj.CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(WbemComputer,
      'root\CIMV2', WbemUser, WbemPassword);
    FWbemQuery := Format(
      'Select * From Win32_PingStatus Where Address=%s And BufferSize=%d And TimeOut=%d',
      [
      QuotedStr(Address), BufferSize, Timeout]);
    FWbemObjectSet := FWMIService.ExecQuery(FWbemQuery, 'WQL', wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, CeltFetched) = 0 do
    begin
      Result := longint(FWbemObject.Properties_.Item('StatusCode').Value);
      FWbemObject := Unassigned;
    end;
  finally
    CoUninitialize;
  end;
end;

procedure TPingThread.Execute;
var
  i, k: integer;
  it: TListItem;
begin
  it := fMain.ListView1.Items[Id];
  it.SubItems[1] := it.SubItems[1] + '*';
  k := 0;
  for i := 1 to Retries do
  begin
    it.SubItems[1] := it.SubItems[1] + '*';
    if WMIPing(Computer, Buffer, Timeout) = 0 then
      Inc(k);
  end;
  it.SubItems[1] := Format('%d of %d ping ok', [k, Retries]);
  Dec(ActiveThreads);
  fMain.StatusBar1.Panels[0].Text := Format('%d active threads', [ActiveThreads]);
  fMain.StatusBar1.Panels[1].Text :=
    Format('%f seconds elapsed', [(GetTickCount - StartTime) / 1000]);
end;

procedure TfMain.ToolButton1Click(Sender: TObject);
var
  sL: TStringList;
  s: string;
begin
  if OpenDialog1.Execute then
  begin
    sL := TStringList.Create;
    sL.LoadFromFile(OpenDialog1.FileName);
    ListView1.Items.Clear;
    for s in sL do
      if Length(s) > 0 then
        with ListView1.Items.Add do
        begin
          Caption := IntToStr(ListView1.Items.Count);
          SubItems.Add(s);
          SubItems.Add('');
        end;
    sL.Free;
    ToolButton2.Enabled := ListView1.Items.Count > 0;
  end;
end;

procedure TfMain.FormCreate(Sender: TObject);
begin
  PageControl1.TabIndex := 0;
end;

procedure TfMain.ToolButton2Click(Sender: TObject);
var
  t: TPingThread;
  i, MaxThreads: integer;
begin
  StartTime := Windows.GetTickCount;
  ActiveThreads := 0;
  MaxThreads := seMaxThreads.Value;
  i := 0;
  StatusBar1.Panels[2].Text := 'working...';
  while i < ListView1.Items.Count do
  begin
    Application.ProcessMessages;
    //Sleep(2);
    if ActiveThreads < MaxThreads then
    begin
      t := TPingThread.Create(True);
      Inc(ActiveThreads);
      Inc(i);
      StatusBar1.Panels[0].Text := Format('%d active threads', [ActiveThreads]);
      StatusBar1.Panels[1].Text :=
        Format('%f seconds elapsed', [(GetTickCount - StartTime) / 1000]);
      t.Id := i - 1;
      t.Buffer := seBuffer.Value;
      t.Timeout := seTimeout.Value;
      t.Retries := seRetries.Value;
      t.Computer := ListView1.Items[i - 1].SubItems[0];
      t.FreeOnTerminate := True;
      t.Start;
      ListView1.Items[i - 1].SubItems[1] := '*';
      Listview1.items[i - 1].MakeVisible(False);
    end;
  end;

  while ActiveThreads > 0 do ;

  StatusBar1.Panels[2].Text := 'done';
end;

procedure TfMain.ToolButton3Click(Sender: TObject);
begin
  if MessageDlg('Reset values to default?', mtConfirmation, [mbYes, mbNo], 0) =
    mrYes then
  begin
    seBuffer.Value := 32;
    seTimeout.Value := 500;
    seRetries.Value := 4;
    seMaxThreads.Value := 30;
  end;
end;

procedure TfMain.ToolButton5Click(Sender: TObject);
var
  si: Windows._SYSTEM_INFO;
begin
  GetSystemInfo(si);
  ShowMessageFmt('The processor of this computer has %d cores.',
    [si.dwNumberOfProcessors]);
end;

end.
