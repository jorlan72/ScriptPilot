unit main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Menus, Vcl.ComCtrls,
  Vcl.StdCtrls, Vcl.WinXPanels, Vcl.Grids, Vcl.ValEdit, snippets, jpeg,
  Vcl.Imaging.pngimage, FlexCel.Core, FlexCel.XlsAdapter, Engine, Vcl.AppEvnts,
  ScSSHChannel, ScSSHClient, ScBridge, System.UITypes, ipwcore, ipwtypes, adddata,
  addtask, ipwping, OverbyteIcsWndControl, OverbyteIcsTnCnx, Vcl.Themes, ShellAPI,
  ipwsyslog, variable, System.Generics.Collections;

type
  TFrmMain = class(TForm)
    StatusBar: TStatusBar;
    MainMenu: TMainMenu;
    MenuFile: TMenuItem;
    ProgressBar: TProgressBar;
    CardPanelLeft: TCardPanel;
    CardPanelRight: TCardPanel;
    CardLog: TCard;
    MemoLog: TMemo;
    CardMenu: TCard;
    DropDownProject: TComboBox;
    DropDownTask: TComboBox;
    GroupBoxProject: TGroupBox;
    GroupBoxTask: TGroupBox;
    GroupBoxDataSource: TGroupBox;
    DropDownData: TComboBox;
    GroupBoxInput: TGroupBox;
    ValueListEditorInput: TValueListEditor;
    GroupBoxLog: TGroupBox;
    GroupBoxActions: TGroupBox;
    ButtonRun: TButton;
    ButtonStop: TButton;
    ImageLogo: TImage;
    PanelLogo: TPanel;
    LabelAppName: TLabel;
    LabelVersion: TLabel;
    LabelCodeRed: TLabel;
    ButtonContinue: TButton;
    CardVariables: TCard;
    ValueListEditorVariables: TValueListEditor;
    MenuExit: TMenuItem;
    MenuView: TMenuItem;
    MenuViewLog: TMenuItem;
    MenuViewVariables: TMenuItem;
    CardVariableInput: TCard;
    CardMessageImage: TCard;
    CardWelcomeLeft: TCard;
    CardWelcomeRight: TCard;
    ButtonStart: TButton;
    ImageFrontPageLogo: TImage;
    LabelFrontPageCodeRed: TLabel;
    MenuEdit: TMenuItem;
    LabelFrontPageBuild: TLabel;
    ButtonDocumentaion: TButton;
    LabelTaskInformation: TLabel;
    TimerOpening: TTimer;
    LabelCounter: TLabel;
    LabelMessageImage: TLabel;
    MessageImage: TImage;
    ButtonMessageImageContinue: TButton;
    LabelAskForInput: TLabel;
    LabelFP05: TLabel;
    LabelFP04: TLabel;
    LabelFP03: TLabel;
    LabelFP02: TLabel;
    LabelFP01: TLabel;
    ApplicationEvents: TApplicationEvents;
    sshclient01: TScSSHClient;
    sshshell01: TScSSHShell;
    ScMemoryStorage1: TScMemoryStorage;
    TimerSSH: TTimer;
    MemoContent: TMemo;
    CardVersionHistory: TCard;
    GroupBoxVersionHistory: TGroupBox;
    MemoVersionHistory: TMemo;
    MenuVersionLog: TMenuItem;
    N1: TMenuItem;
    LabelScriptActiveRow: TLabel;
    Ping: TipwPing;
    MenuOpenProjectFolder: TMenuItem;
    MenuEditDataSource: TMenuItem;
    MenuEditTaskExcelFile: TMenuItem;
    MenuCreate: TMenuItem;
    MenuCreateNewProject: TMenuItem;
    MenuCreateNewTask: TMenuItem;
    MenuHelp: TMenuItem;
    MenuDocumentation: TMenuItem;
    GroupBoxInformation: TGroupBox;
    LabelInformation: TLabel;
    ImageInformation: TImage;
    MenuRefreshProject: TMenuItem;
    MenuProjectDocuments: TMenuItem;
    MenuClearLog: TMenuItem;
    N5: TMenuItem;
    N4: TMenuItem;
    MenuRun: TMenuItem;
    MenuStop: TMenuItem;
    CBWelcomeStartup: TCheckBox;
    TelnetClient01: TTnCnx;
    SysLog: TipwSysLog;
    CardSelector: TCard;
    ComboBoxSelector: TComboBox;
    ButtonSelectorContinue: TButton;
    LabelSelector: TLabel;
    MenuFloatingVariables: TMenuItem;
    N2: TMenuItem;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure DropDownProjectSelect(Sender: TObject);
    procedure DropDownDataSelect(Sender: TObject);
    procedure ButtonRunClick(Sender: TObject);
    procedure MenuExitClick(Sender: TObject);
    procedure MenuViewLogClick(Sender: TObject);
    procedure MenuViewVariablesClick(Sender: TObject);
    procedure ButtonStartClick(Sender: TObject);
    procedure DropDownTaskSelect(Sender: TObject);
    procedure TimerOpeningTimer(Sender: TObject);
    procedure ButtonStopClick(Sender: TObject);
    procedure ButtonMessageImageContinueClick(Sender: TObject);
    procedure ButtonContinueClick(Sender: TObject);
    procedure ApplicationEventsHint(Sender: TObject);
    procedure sshclient01ServerKeyValidate(Sender: TObject; NewServerKey: TScKey;
      var Accept: Boolean);
    procedure sshclient01Banner(Sender: TObject; const Banner: string);
    procedure sshshell01AsyncReceive(Sender: TObject);
    procedure TimerSSHTimer(Sender: TObject);
    procedure MenuVersionLogClick(Sender: TObject);
    procedure ButtonDocumentaionClick(Sender: TObject);
    procedure PingResponse(Sender: TObject; RequestId: Integer;
      const ResponseSource, ResponseStatus: string; ResponseTime: Integer);
    procedure MenuOpenProjectFolderClick(Sender: TObject);
    procedure MenuEditDataSourceClick(Sender: TObject);
    procedure MenuEditTaskExcelFileClick(Sender: TObject);
    procedure MenuCreateNewProjectClick(Sender: TObject);
    procedure MenuRefreshProjectClick(Sender: TObject);
    procedure MenuCreateNewTaskClick(Sender: TObject);
    procedure MenuClearLogClick(Sender: TObject);
    procedure MenuDocumentationClick(Sender: TObject);
    procedure CBWelcomeStartupClick(Sender: TObject);
    procedure MenuItemClick(Sender: TObject);
    procedure TelnetClient01DataAvailable(Sender: TTnCnx; Buffer: Pointer;
      Len: Integer);
    procedure MenuCalypsoClick(Sender: TObject);
    procedure ButtonSelectorContinueClick(Sender: TObject);
    procedure MenuFloatingVariablesClick(Sender: TObject);
    procedure ValueListEditorVariablesStringsChange(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private

  public
    SSHClientList: TObjectList<TScSSHClient>;
    SSHShellList: TObjectList<TScSSHShell>;
  end;

type
  Tdata = Record
   InputDescription: variant;
   VariableName: variant;
   InitialValue: variant;
  End;

  TDataArray = array of TData;

var
  FrmMain: TFrmMain;
  ExcelCommand : variant;
  ExcelParam01 : variant;
  ExcelParam02 : variant;
  ExcelParam03 : variant;
  ExcelString01, ExcelString02, ExcelString03 : string;
  CurrentRow : Integer;
  TimerCounter : Integer;
  StopPressed : boolean;
  ContinuePressed : boolean;
  InputArray : TDataArray;
  ExternalProcess: TExternalProcess;
  Verbose : boolean;
  Timeout : integer;
  MaxTimeout : Integer;
  ChillTime : Integer;
  FeedbackReceived : boolean;
  LastVerboseStart : Integer;
  sshinputstring : string;
  telnetinputstring : string;
  StringsToRemove : TStringList;
  QuitApplication : Boolean;
  DoStringReplace : Boolean;
  GotReplyFromPing : Boolean;
  PingHost : String;
  WillTrim : Boolean;
  StartTime: TDateTime;
  WillLog : Boolean;
  NewProject : Boolean;
  NewProjectText : String;
  taskfilename : string;
  datafilename : string;
  menutag : string;
  MenuAllreadyAdded : Boolean;
  NumberOfCommands : string;
  PickList : tStringlist;
  ipv4, ipv6, switchname : string;
  ipv4gw, ipv6gw : string;
  switchstartipv4, switchstartipv6 : string;
  ipv4temp, ipv6temp : integer;



implementation


{$R *.dfm}


procedure TFrmMain.MenuItemClick(Sender: TObject);
begin
  if Sender is TMenuItem then
    ShellOpen(TMenuItem(Sender).Hint);
  LogItStamp('Opening project document url - ' + TMenuItem(Sender).Hint,0);
end;

procedure StartTimer;
begin
  StartTime := Now;
end;

function StopTimer: string;
var
  ElapsedTime: TDateTime;
begin
  ElapsedTime := Now - StartTime;
  Result := FormatDateTime('hh:nn:ss.zzz', ElapsedTime);
end;

procedure TFrmMain.ApplicationEventsHint(Sender: TObject);
begin
 frmMain.StatusBar.SimpleText := application.hint;
end;

procedure TFrmMain.ButtonContinueClick(Sender: TObject);
var
 AnyBlanks : boolean;
 Counter : integer;
begin
  AnyBlanks := false;
      for Counter := 0 to frmMain.ValueListEditorInput.Strings.Count -1 do begin
        if frmMain.ValueListEditorInput.strings.ValueFromIndex[Counter] = '' then AnyBlanks := true
      end;
   if AnyBlanks then ShowMessage('          Please input all values!        ') else ContinuePressed := true;
   if AnyBlanks then LogItStamp('Blank input not allowed. Altering user with a MessageBox',0);
end;


procedure TFrmMain.ButtonDocumentaionClick(Sender: TObject);
begin
 if FileExists(extractfilepath(application.ExeName) + '\ScriptPilot Documentation.pdf') then ShellOpen(extractfilepath(application.ExeName) + '\ScriptPilot Documentation.pdf');
end;

procedure TFrmMain.ButtonMessageImageContinueClick(Sender: TObject);
begin
ContinuePressed := true;
end;

procedure TFrmMain.ButtonRunClick(Sender: TObject);
var
 xls : TXlsFile;
 Counter : integer;
 FileName : string;
 VariableCounter : integer;
 Cell : TCellAddress;
 LabelName : variant;
 row,col : integer;
begin
//frmMain.Visible := false;
StartTimer;
Progressbar.Position := 0;
frmMain.LabelScriptActiveRow.Visible := true;
TerminateExternalProcessFlag := False;
frmMain.TimerSSH.Enabled := false;
Verbose := false;
frmMain.LabelScriptActiveRow.Caption := '';
LogIt('',0);
frmMain.GroupBoxProject.Enabled := false;
frmMain.GroupBoxDataSource.Enabled := false;
frmMain.GroupBoxTask.Enabled := false;
frmMain.ButtonRun.Enabled := false;
frmMain.MenuRun.Enabled := false;
frmMain.ButtonStop.Enabled := true;
frmMain.MenuStop.Enabled := True;
frmMain.MenuFile.Visible := false;
frmMain.MenuView.Visible := false;
frmMain.MenuEdit.Visible := false;
frmMain.MenuCreate.Visible := false;
// Disable shortcuts begin //
frmMain.MenuOpenProjectFolder.Enabled := false;
frmMain.MenuExit.Enabled := false;
frmMain.MenuEditDataSource.Enabled := false;
frmMain.MenuEditTaskExcelFile.Enabled := false;
frmMain.MenuRefreshProject.Enabled := false;
frmMain.MenuClearLog.Enabled := false;
// Disable shortcuts end //
frmMain.MenuProjectDocuments.Visible := false;
frmMain.MenuHelp.Visible := false;
frmMain.LabelScriptActiveRow.Visible := true;
frmMain.StatusBar.simpletext := 'Script running';
frmMain.CardPanelRight.ActiveCardIndex := 1;
frmMain.CardPanelLeft.ActiveCardIndex := 1;
frmMain.MenuViewLog.click;
StopPressed := false;
WillTrim := false;
frmMain.LabelInformation.Caption := '';
LogItStamp('User hit the Run button on selected script ' + frmMain.DropDownTask.Text,0);
LogItStamp('Disabling menus, drop-down boxes and the Run button. Enabling the Stop button',1);
LogIt('************************************************************************************************************************',0);
LogIt(Uppercase(frmmain.DropDownTask.Text) + ' : SCRIPT STARTING -  USING DATA SOURCE "' + Uppercase(frmMain.DropDownData.text) + '"',1);
UpdateSettings;
sshclient01.Timeout := chilltime;
sshshell01.Timeout := chilltime;
UpdateVariables;
  xls := TXlsFile.Create;
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.xlsx';
  CopyFileToTemp(filename);
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.tmp';
      if SheetExists(filename,'Script') then begin
          xls.open(filename);
          xls.ActiveSheetByName := 'Script';
          LogItStamp('Found Script sheet. Processing commands from row 2',0);
          // fetch labels //
          Cell := TCellAddress.Empty;
             repeat
               Cell := xls.Find('#label#', txlscellrange.Null, cell, false, true, false, true);
                 if cell.HasValue then
                 begin
                   row := cell.row;
                   col := cell.col;
                   LabelName := xls.GetCellValue(row,col+1);
                   LogItStamp('Found label "' + labelname + '" at row ' + inttostr(row) + '. Adding to variable table', 0);
                   frmmain.ValueListEditorVariables.InsertRow(LabelName, inttostr(row), true);
                 end;
             until Cell.IsNull;
          CurrentRow := 2;
          Counter := 2;
                while not xls.GetCellValue(counter,1).IsEmpty do begin
                DoStringReplace := true;
                ExcelCommand := xls.GetCellValue(Counter,1);
                ExcelParam01 := '';
                ExcelParam02 := '';
                ExcelParam03 := '';
                 if not xls.GetCellValue(Counter,2).IsEmpty then ExcelParam01 := xls.GetCellValue(Counter,2);
                 if not xls.GetCellValue(Counter,3).IsEmpty then ExcelParam02 := xls.GetCellValue(Counter,3);
                 if not xls.GetCellValue(Counter,4).IsEmpty then ExcelParam03 := xls.GetCellValue(Counter,4);
                // *****************  stringreplace paramstring variables here before processing
                  if vartostr(ExcelParam01) <> '' then begin
                  // Commands to not stringreplace under //
                  if LowerCase(ExcelCommand) = 'variablecreate' then DoStringReplace := false;
                  if LowerCase(ExcelCommand) = 'variableset' then DoStringReplace := false;
                  if LowerCase(ExcelCommand) = 'askforinputtable' then DoStringReplace := false;
                  if LowerCase(ExcelCommand) = 'askforinput' then DoStringReplace := false;
                  if LowerCase(ExcelCommand) = 'fetchlogtovariable' then DoStringReplace := false;
                  if LowerCase(ExcelCommand) = 'variableincrease' then DoStringReplace := false;
                  if LowerCase(ExcelCommand) = 'variabledecrease' then DoStringReplace := false;
                  if LowerCase(ExcelCommand) = 'gotoifvariableequals' then DoStringReplace := false;
                  // Commands to not stringreplace above //
                     for VariableCounter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do begin
                       ExcelString01 := vartostr(ExcelParam01);
                       if DoStringReplace then ExcelParam01 := StringReplace(ExcelString01, frmMain.ValueListEditorVariables.Strings.KeyNames[VariableCounter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[VariableCounter], [rfReplaceAll]);
                       Application.ProcessMessages;
                     end;
                   DoStringReplace := true;
                  end;
                  if vartostr(ExcelParam02) <> '' then begin
                     for VariableCounter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do begin
                       ExcelString02 := vartostr(ExcelParam02);
                       if DoStringReplace then ExcelParam02 := StringReplace(ExcelString02, frmMain.ValueListEditorVariables.Strings.KeyNames[VariableCounter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[VariableCounter], [rfReplaceAll]);
                       Application.ProcessMessages;
                     end;
                  end;
                  if vartostr(ExcelParam03) <> '' then begin
                     for VariableCounter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do begin
                       ExcelString03 := vartostr(ExcelParam03);
                       if DoStringReplace then ExcelParam03 := StringReplace(ExcelString03, frmMain.ValueListEditorVariables.Strings.KeyNames[VariableCounter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[VariableCounter], [rfReplaceAll]);
                       Application.ProcessMessages;
                     end;
                  end;
                CurrentRow := Counter;
                LabelScriptActiveRow.caption := 'Script row ' + inttostr(CurrentRow) + ' : ' + vartostr(ExcelCommand) + ' : ' + vartostr(ExcelParam01) + ' : ' + vartostr(ExcelParam02) + ' : ' + vartostr(ExcelParam03);
                ProcessCommands;
                //If CurrentRow is changed by the script//
                Counter := CurrentRow;
                inc(Counter);
                if StopPressed then Counter := 9000;
                if QuitApplication then frmMain.Close;
                end;
      end else begin
        LogItStamp('ERROR: Script sheet not found. No commands to execute',0);
      end;
  xls.free;
// If the Stop button was pressed during execution
 if StopPressed then begin
   LogItStamp('The user stopped to script before it could finish at script row ' + inttostr(CurrentRow),0);
 end;
// The following is run when the entire script is processed
WillLog := true;
if not stoppressed then LabelScriptActiveRow.Caption := 'Script ended at row ' + inttostr(currentrow) else
                        LabelScriptActiveRow.Caption := 'Script aborted at row ' + inttostr(currentrow);
if not stoppressed then LogItStamp('Empty cell encountered at row ' + inttostr(CurrentRow + 1) + '. Assuming end of script',0);
if not stoppressed then LogItStamp('The selected script "' + frmMain.DropDownTask.Text + '" has completed - Elapsed time : ' + StopTimer,1)
else LogItStamp('The selected script ' + frmMain.DropDownTask.Text + ' was aborted before it could finish.',1);
if not stoppressed then LogIt(Uppercase(frmmain.DropDownTask.Text) + ' : SCRIPT FINISHED - USING DATA SOURCE "' + Uppercase(frmMain.DropDownData.text) + '"',0)
else LogIt(Uppercase(frmmain.DropDownTask.Text) + ' : SCRIPT ABORTED - USING DATA SOURCE "' + Uppercase(frmMain.DropDownData.text) + '"',0);
LogIt('************************************************************************************************************************',1);
//frmMain.Visible := false;
Progressbar.Position := 0;
LogItStamp('Enabling menus, drop-down boxes and the Run button. Disabling the Stop button',0);
DeleteFile(filename);
frmMain.GroupBoxProject.Enabled := true;
frmMain.GroupBoxDataSource.Enabled := true;
frmMain.GroupBoxTask.Enabled := true;
frmMain.ButtonRun.Enabled := true;
frmMain.MenuRun.Enabled := true;
frmMain.MenuFile.Visible := true;
frmMain.MenuView.Visible := true;
frmMain.MenuEdit.Visible := true;
frmMain.MenuCreate.Visible := true;
frmMain.MenuProjectDocuments.Visible := true;
frmMain.MenuHelp.Visible := true;
frmMain.ButtonStop.Enabled := false;
frmMain.MenuStop.Enabled := false;
frmMain.MenuViewLog.Click;
frmMain.StatusBar.simpletext := 'Script Done';
// Enable shortcuts begin //
frmMain.MenuOpenProjectFolder.Enabled := true;
frmMain.MenuExit.Enabled := true;
frmMain.MenuEditDataSource.Enabled := true;
frmMain.MenuEditTaskExcelFile.Enabled := true;
frmMain.MenuRefreshProject.Enabled := true;
frmMain.MenuClearLog.Enabled := true;
// Enable shortcuts end //
// Reset variables
frmMain.ValueListEditorVariables.Strings.Clear;
frmmain.ValueListEditorVariables.InsertRow('#DataSourceSelected#', frmMain.DropDownData.text, true);
frmmain.ValueListEditorVariables.InsertRow('#DataSourceIndex#', inttostr(frmMain.DropDownData.ItemIndex + 1), true);
frmmain.ValueListEditorVariables.InsertRow('#DataSourceItemsCount#', inttostr(frmmain.DropDownData.Items.Count), true);
frmmain.ValueListEditorVariables.InsertRow('#ApplicationPath#', extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text, true);
UpdateVariables;
LabelScriptActiveRow.caption := '';
frmMain.GroupBoxInformation.Visible := false;
frmMain.ImageInformation.Picture := nil;
frmMain.LabelScriptActiveRow.Visible := false;
LogItStamp('Ready.',1);
end;

procedure TFrmMain.ButtonSelectorContinueClick(Sender: TObject);
begin
ContinuePressed := true;
end;

procedure TFrmMain.ButtonStartClick(Sender: TObject);
var
 StartProject, StartDataSource, StartTask : Integer;
begin
if not newproject then frmMain.TimerOpening.Enabled := false;
if not newproject then frmMain.Visible := false;
if not newproject then frmMain.CardPanelRight.ActiveCardIndex := 1;
if not newproject then frmMain.CardPanelLeft.ActiveCardIndex := 1;
if not newproject then frmMain.MenuFile.Visible := true;
if not newproject then frmMain.MenuView.Visible := true;
if not newproject then frmMain.MenuEdit.Visible := true;
if not newproject then frmMain.MenuCreate.Visible := true;
if not newproject then frmMain.MenuProjectDocuments.Visible := true;
if not newproject then frmMain.MenuHelp.Visible := true;
if not newproject then frmMain.CardPanelRight.BevelInner := bvLowered;
if not newproject then frmMain.CardPanelRight.BevelKind := bkFlat;
if not newproject then frmMain.CardPanelRight.BevelOuter := bvRaised;
if not newproject then frmMain.CardPanelRight.BevelWidth := 2;
if not newproject then frmMain.CardPanelRight.BorderStyle := bsSingle;
if not newproject then frmMain.CardPanelRight.BorderWidth := 1;
if not newproject then frmMain.BorderStyle := bsSizeable;
if not newproject then frmMain.WindowState := wsMaximized;
if not newproject then Logit('',0);
if not newproject then Logit('        _/_/_/                      _/              _/       _/_/_/    _/  _/              _/      ',0);
if not newproject then Logit('     _/          _/_/_/  _/  _/_/      _/_/_/    _/_/_/_/   _/    _/      _/    _/_/    _/_/_/_/   ',0);
if not newproject then Logit('      _/_/    _/        _/_/      _/  _/    _/    _/       _/_/_/    _/  _/  _/    _/    _/        ',0);
if not newproject then Logit('         _/  _/        _/        _/  _/    _/    _/       _/        _/  _/  _/    _/    _/         ',0);
if not newproject then Logit('  _/_/_/      _/_/_/  _/        _/  _/_/_/        _/_/   _/        _/  _/    _/_/        _/_/      ',0);
if not newproject then Logit('                                   _/                                                             ',0);
if not newproject then Logit('                                  _/                              CodeRed - 2023 - J. Lanesskog',2);
if not newproject then LogIt('Initializing application - Version Build : ' + GetAppVersionStr,0);
if not newproject then LogIt(NumberOfCommands,0);
if not newproject then LogIt('For documentation, hit F1',1);
// Disable features before task is selected
frmMain.ButtonRun.Enabled := false;
frmMain.MenuRun.Enabled := false;
frmMain.ButtonStop.Enabled := false;
frmMain.MenuStop.Enabled := false;
// Populate drop downs
PopulateComboBoxWithSubfolders(extractfilepath(application.ExeName) + 'projects', DropDownProject);
// Add all variables to the application from the Excel Data Source
if frmMain.DropDownProject.ItemIndex <> -1 then InitializeVariablesAndData else
   begin
    LogItStamp('No projects found. Select "Create new Project" from the "Create" menu',0);
   end;
// Add all task files to drop-down
newproject := false;
if frmMain.DropDownData.ItemIndex <> -1 then PopulateComboBoxWithFilesOfType(extractfilepath(application.ExeName) + 'projects\' + frmMain.DropDownProject.Text + '\Tasks', '*.xlsx', DropDownTask);
// Enable the Run button if there is an active task present in the task drop-down
if frmMain.DropDownTask.ItemIndex = 0 then frmMain.ButtonRun.Enabled := true;
if frmMain.DropDownTask.ItemIndex = 0 then frmMain.MenuRun.Enabled := true;
if frmMain.DropDownTask.ItemIndex = 0 then LogItStamp('Tasks found. Run button enabled',0);
// Update settings for current task
if frmMain.DropDownData.ItemIndex <> -1 then UpdateSettings;
// Update version number
frmMain.LabelVersion.Caption := 'Version Build : ' + GetAppVersionStr;
frmMain.MenuOpenProjectFolder.enabled := true;
frmMain.MenuExit.Enabled := true;
frmMain.MenuRun.Enabled := true;
frmMain.MenuEditDataSource.enabled := true;
frmMain.MenuEditTaskExcelFile.enabled := true;
frmMain.MenuRefreshProject.Enabled := true;
frmMain.MenuClearLog.enabled := true;
frmMain.MenuCreateNewProject.enabled := true;
frmMain.MenuCreateNewTask.enabled := true;
frmMain.MenuFloatingVariables.enabled := true;
frmMain.MenuDocumentation.enabled := true;
frmMain.Visible := true;
// Check if application was started with any parameters
  if trystrtoint(ParamStr(2), StartProject) then
    begin
    LogItStamp('Application startup parameter #2 = ' + inttostr(StartProject),0);
    if startProject < 1 then exit;
      if StartProject <= frmMain.DropDownProject.Items.count then
        begin
         frmMain.DropDownProject.ItemIndex := StartProject - 1;
         DropDownProjectSelect(frmMain.DropDownProject);
        end;
    end;
  if trystrtoint(ParamStr(3), StartDataSource) then
    begin
    LogItStamp('Application startup parameter #3 = ' + inttostr(StartDataSource),0);
    if StartDataSource < 1 then exit;
      if StartDataSource <= frmMain.DropDownData.Items.count then
       begin
        frmMain.DropDownData.ItemIndex := StartDataSource - 1;
        DropDownDataSelect(frmMain.DropDownData);
       end;
    end;
  if trystrtoint(ParamStr(4), StartTask) then
    begin
    LogItStamp('Application startup parameter #4 = ' + inttostr(StartTask),0);
    if startTask < 1 then exit;
     if StartTask <= frmMain.DropDownTask.Items.count then
      begin
       frmMain.DropDownTask.ItemIndex := StartTask - 1;
       DropDownTaskSelect(frmMain.DropDownTask);
      end;
    end;
  if lowercase(ParamStr(5)) = 'run' then
   begin
    LogItStamp('Application startup parameter #5 = ' + paramstr(5),0);
    LogItStamp('Script will autostart now',0);
    frmMain.ButtonRun.Click;
   end;
LogItStamp('Ready.',1);
end;

procedure TFrmMain.ButtonStopClick(Sender: TObject);
begin
LogItStamp('User clicked the Stop button. Script will stop execution..',0);
StopPressed := true;
TerminateExternalProcessFlag := True;
end;

procedure TFrmMain.CBWelcomeStartupClick(Sender: TObject);
begin
 if frmMain.CBWelcomeStartup.Checked then
    begin
     if WriteRegistryString(HKEY_CURRENT_USER, 'Software\CodeRed\ScriptPilot', 'LingerWelcome', 'True') then
      LogItStamp('INFO: Welcome screen will linger for 30 seconds at next startup',1)
    end else
    begin
     if WriteRegistryString(HKEY_CURRENT_USER, 'Software\CodeRed\ScriptPilot', 'LingerWelcome', 'False') then
      LogItStamp('INFO: Welcome screen will NOT linger at next startup',1)
    end;
end;

procedure TFrmMain.DropDownDataSelect(Sender: TObject);
begin
if frmMain.DropDownTask.ItemIndex = -1 then frmMain.ButtonRun.Enabled := false;
if frmMain.DropDownTask.ItemIndex = -1 then frmMain.MenuRun.Enabled := false;
UpdateVariables;
end;

procedure TFrmMain.DropDownProjectSelect(Sender: TObject);
begin
LogIt('',0);
frmMain.DropDownData.Clear;
frmMain.DropDownTask.Clear;
MenuAllreadyAdded := false;
InitializeVariablesAndData;
if frmMain.DropDownData.ItemIndex <> -1 then PopulateComboBoxWithFilesOfType(extractfilepath(application.ExeName) + 'projects\' + frmMain.DropDownProject.Text + '\Tasks', '*.xlsx', DropDownTask);
if frmMain.DropDownTask.ItemIndex = -1 then frmMain.ButtonRun.Enabled := false else frmMain.ButtonRun.Enabled := true;
if frmMain.DropDownTask.ItemIndex = -1 then frmMain.MenuRun.Enabled := false else frmMain.MenuRun.Enabled := true;
if frmMain.DropDownTask.ItemIndex = -1 then LogItStamp('Tasks disabled. Run button disabled',0) else LogItStamp('Tasks found. Run button enabled',0);
if frmmain.DropDownData.items.Count = 0 then frmmain.DropDownData.Enabled := false else frmmain.DropDownData.Enabled := true;
if frmmain.DropDownTask.items.Count = 0 then frmmain.DropDownTask.Enabled := false else frmmain.DropDownTask.Enabled := true;
if frmmain.DropDownTask.items.Count > 0 then UpdateSettings;
end;

procedure TFrmMain.DropDownTaskSelect(Sender: TObject);
begin
 UpdateSettings;
end;

procedure TFrmMain.MenuCalypsoClick(Sender: TObject);
begin
  TStyleManager.TrySetStyle('Calypso');
   if WriteRegistryString(HKEY_CURRENT_USER, 'Software\CodeRed\ScriptPilot', 'Skin', 'Calypso') then
   LogItStamp('INFO: Calypso set as application skin. Restarting application',0)
end;

procedure TFrmMain.MenuClearLogClick(Sender: TObject);
begin
frmMain.CardPanelRight.ActiveCardIndex := 1;
frmMain.MemoLog.Clear;
LogItStamp('Ready.',0);
end;

procedure TFrmMain.MenuCreateNewProjectClick(Sender: TObject);
var
 inputreply : string;
 sourcedir, destdir : string;
begin
// frmMain.CardPanelRight.ActiveCardIndex := 1;
 LogItStamp('Creating new project. Asking user for project name.',0);
 if Inputquery('Create a new project', 'Name of the new project :', inputreply) then
 begin
     if inputreply = '' then
        begin
         LogItStamp('ERROR: Project name cannot be blank. Will NOT create the new project.',0);
         exit;
        end;
     if FindIndexByText(frmMain.DropDownProject, inputreply) <> -1 then
        begin
         LogItStamp('ERROR: Project name already exists. Will NOT create the new project.',0);
         exit;
        end;
      if CreateDirectoryIfNotExists(extractfilepath(application.ExeName) + '\Projects\') then
        begin
         LogItStamp('Creating "Projects" directory.',0);
        end;
      if CreateDirectoryIfNotExists(extractfilepath(application.ExeName) + '\Template\') then
        begin
         LogItStamp('Creating "Template" directory.',0);
        end;
      if CreateDirectoryIfNotExists(extractfilepath(application.ExeName) + '\Template\Data') then
        begin
         LogItStamp('Creating "Template\Data" directory and bulding initial data.xlsx file.',0);
         datafilename := extractfilepath(application.ExeName) + '\Template\Data\Data.xlsx';
         CreateAndSaveDataFile;
        end;
      if CreateDirectoryIfNotExists(extractfilepath(application.ExeName) + '\Template\Tasks') then
        begin
         LogItStamp('Creating "Template\Tasks" directory and bulding initial task.xlsx file.',0);
         taskfilename := extractfilepath(application.ExeName) + '\Template\Tasks\New task.xlsx';
         CreateAndSaveTaskFile;
        end;
  LogItStamp('Creating project "' + inputreply + '"',0);
  LogItStamp('Adding template files.',0);
  sourcedir := extractfilepath(application.ExeName) + '\Template\';
  destdir := extractfilepath(application.ExeName) + '\Projects\' + inputreply;
    if CopyDirectory(sourcedir, destdir) then
        begin
         LogItStamp('Project "' + inputreply + '" created',0);
         LogItStamp('INFO: To delete or rename a project, do so manually in the "projects" folder',0);
         LogItStamp('INFO: Any direct folder level changes will be available after hitting F5 (refresh) or by restarting the application',0);
         newproject := true;
         newprojecttext := inputreply;
         frmMain.DropDownData.Clear;
         frmMain.DropDownTask.Clear;
         frmMain.ButtonStart.Click;
         if frmmain.DropDownData.items.Count = 0 then frmmain.DropDownData.Enabled := false else frmmain.DropDownData.Enabled := true;
         if frmmain.DropDownTask.items.Count = 0 then frmmain.DropDownTask.Enabled := false else frmmain.DropDownTask.Enabled := true;
        end else begin
         LogItStamp('ERROR: Creating new project failed. Check source template dir or file permissions',0);
        end;
 end else
     begin
      LogItStamp('Creating new project. Cancelled by user.',0);
     end;
end;

procedure TFrmMain.MenuCreateNewTaskClick(Sender: TObject);
var
 inputreply : string;
 sourcefile, destdir : string;
begin
 if frmMain.DropDownData.ItemIndex = -1 then
    begin
     LogItStamp('ERROR: No project to add tasks to. Create a project first.',0);
     exit;
    end;
LogItStamp('Creating a new task. Asking user for task name.',0);
if Inputquery('Create a new task', 'Name of the new task :', inputreply) then
   begin
     if inputreply = '' then
        begin
         LogItStamp('ERROR: Task name cannot be blank. Will NOT create the new task.',0);
         exit;
        end;
     if FindIndexByText(frmMain.DropDownTask, inputreply) <> -1 then
        begin
         LogItStamp('ERROR: Task name already exists. Will NOT create the new task.',0);
         exit;
        end;
     sourcefile := extractfilepath(application.ExeName) + '\Template\Tasks\New task.xlsx';
     destdir := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\';
      if CopyFileWithNewName(sourcefile, destdir, inputreply + '.xlsx') then
        begin
          LogItStamp('Creating task "' + inputreply + '"',0);
          LogItStamp('Opening new task for editing.',0);
          ShellOpen(extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + inputreply + '.xlsx');
          LogItStamp('INFO: To delete or rename a task, do so manually in the "projects" folder',0);
          LogItStamp('INFO: Any direct file level changes will be available after hitting F5 (refresh) or by restarting the application',0);
          LogItStamp('INFO: Added material, like images for the script, should be placed in the projects "Data" folder.',0);
          PopulateComboBoxWithFilesOfType(extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\', '*.xlsx', frmMain.DropDownTask);
          frmMain.DropDownTask.ItemIndex := FindIndexByText(frmMain.DropDownTask, inputreply);
          updatesettings;
        end;
   end else
         begin
          LogItStamp('Creating new task. Cancelled by user.',0);
         end;
end;

procedure TFrmMain.MenuDocumentationClick(Sender: TObject);
begin
 LogItStamp('Opening the documentation file',0);
 frmMain.ButtonDocumentaion.Click;
end;

procedure TFrmMain.MenuEditDataSourceClick(Sender: TObject);
var
 filename : string;
begin
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + 'data.xlsx';
  if fileexists(filename) then ShellOpen(filename) else LogItStamp('File could not be found',0);
  if fileexists(filename) then LogItStamp('Opening project data file - ' + filename,0);
end;

procedure TFrmMain.MenuEditTaskExcelFileClick(Sender: TObject);
var
 filename : string;
begin
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.xlsx';
  if fileexists(filename) then ShellOpen(filename) else LogItStamp('File could not be found',0);
  if fileexists(filename) then LogItStamp('Opening current task file - ' + filename,0);
end;

procedure TFrmMain.MenuExitClick(Sender: TObject);
begin
 Close;
end;

procedure TFrmMain.MenuFloatingVariablesClick(Sender: TObject);
begin
if frmVariable.visible = true then frmVariable.visible := false else frmVariable.visible := true;
end;

procedure TFrmMain.MenuOpenProjectFolderClick(Sender: TObject);
var
 folder : string;
begin
  folder := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text;
  ShellOpen(folder);
end;

procedure TFrmMain.MenuRefreshProjectClick(Sender: TObject);
begin
frmMain.CardPanelRight.ActiveCardIndex := 1;
LogIt('',1);
LogItStamp('Refreshing all projects, data and tasks',1);
frmMain.DropDownData.Clear;
frmMain.DropDownTask.Clear;
frmMain.DropDownProject.Clear;
frmMain.ButtonStart.Click;
if frmmain.DropDownData.items.Count = 0 then frmmain.DropDownData.Enabled := false else frmmain.DropDownData.Enabled := true;
if frmmain.DropDownTask.items.Count = 0 then frmmain.DropDownTask.Enabled := false else frmmain.DropDownTask.Enabled := true;
end;

procedure TFrmMain.MenuVersionLogClick(Sender: TObject);
begin
 frmMain.CardPanelRight.ActiveCardIndex := 5;
end;

procedure TFrmMain.MenuViewLogClick(Sender: TObject);
begin
 frmMain.CardPanelRight.ActiveCardIndex := 1;
end;

procedure TFrmMain.MenuViewVariablesClick(Sender: TObject);
begin
 frmMain.CardPanelRight.ActiveCardIndex := 2;
end;

procedure TFrmMain.PingResponse(Sender: TObject; RequestId: Integer;
  const ResponseSource, ResponseStatus: string; ResponseTime: Integer);
begin
  if responseSource = PingHost then if responsestatus = 'OK' then GotReplyFromPing := true else GotReplyFromPing := false;
end;

procedure TFrmMain.sshclient01Banner(Sender: TObject; const Banner: string);
begin
LogItStamp('Received welcome banner from SSH host',0);
end;

procedure TFrmMain.sshclient01ServerKeyValidate(Sender: TObject;
  NewServerKey: TScKey; var Accept: Boolean);
begin
 accept := true;
end;

procedure TFrmMain.sshshell01AsyncReceive(Sender: TObject);
var
Counter : Integer;
sshSender: TScSSHShell;
begin
sshSender := Sender as TScSSHShell;
sshinputstring := sshSender.ReadString;
  if Verbose then begin
   MemoContent.Clear;
   MemoContent.Text := sshinputstring;
       if WillTrim then CleanMemoText(MemoContent, StringsToRemove);
    for Counter := 0 to MemoContent.Lines.Count - 1 do begin
     MemoLog.Lines.Add(MemoContent.Lines.Strings[Counter])
    end;
  end;
FeedbackReceived := true;

{sshinputstring := sshshell01.ReadString;
  if Verbose then begin
   MemoContent.Clear;
   MemoContent.Text := sshinputstring;
       if WillTrim then CleanMemoText(MemoContent, StringsToRemove);
    for Counter := 0 to MemoContent.Lines.Count - 1 do begin
     MemoLog.Lines.Add(MemoContent.Lines.Strings[Counter])
    end;
  end;
FeedbackReceived := true;}
end;

procedure TFrmMain.TelnetClient01DataAvailable(Sender: TTnCnx; Buffer: Pointer;
  Len: Integer);
var
Counter : Integer;
Text: AnsiString;
begin
SetString(Text, PAnsiChar(Buffer), Len);
sshinputstring := Text;
  if Verbose then begin
   MemoContent.Clear;
   MemoContent.Text := sshinputstring;
       if WillTrim then CleanMemoText(MemoContent, StringsToRemove);
    for Counter := 0 to MemoContent.Lines.Count - 1 do begin
     MemoLog.Lines.Add(MemoContent.Lines.Strings[Counter])
    end;
  end;
FeedbackReceived := true;
end;

procedure TFrmMain.TimerOpeningTimer(Sender: TObject);
begin
dec(TimerCounter);
frmMain.LabelCounter.Caption := 'This page will close automatically in ' + inttostr(TimerCounter) + ' seconds';
if TimerCounter = 0 then frmMain.ButtonStart.Click;
end;

procedure TFrmMain.TimerSSHTimer(Sender: TObject);
begin
 application.ProcessMessages;
end;

procedure TFrmMain.ValueListEditorVariablesStringsChange(Sender: TObject);
begin
 CopyValueListEditor(frmMain.ValueListEditorVariables, frmVariable.ValueListLiveVariables);
end;

procedure TFrmMain.FormCreate(Sender: TObject);
var
 RegKeyLingerValue : string;
begin
frmMain.Visible := false;
frmMain.ScaleForCurrentDpi;
RegKeyLingerValue := ReadRegistryString(HKEY_CURRENT_USER,'Software\CodeRed\ScriptPilot', 'LingerWelcome', 'True');
if RegKeyLingerValue = 'True' then frmMain.CBWelcomeStartup.Checked := true else frmMain.CBWelcomeStartup.Checked := false;
QuitApplication := false;
WillLog := true;
// Initialize data structures //
{SSHClientList := TObjectList<TScSSHClient>.Create;
SSHShellList := TObjectList<TScSSHShell>.Create;}
NewProject := false;
frmMain.BorderStyle := bsNone;
frmMain.CardPanelRight.ActiveCardIndex := 0;
frmMain.CardPanelLeft.ActiveCardIndex := 0;
frmMain.MenuFile.Visible := false;
frmMain.MenuView.Visible := false;
frmMain.MenuEdit.Visible := false;
frmMain.MenuCreate.Visible := false;
frmMain.MenuProjectDocuments.Visible := false;
frmMain.MenuHelp.Visible := false;
frmMain.MenuOpenProjectFolder.enabled := false;
frmMain.MenuExit.Enabled := false;
frmMain.MenuRun.Enabled := false;
frmMain.MenuEditDataSource.enabled := false;
frmMain.MenuEditTaskExcelFile.enabled := false;
frmMain.MenuRefreshProject.Enabled := false;
frmMain.MenuClearLog.enabled := false;
frmMain.MenuCreateNewProject.enabled := false;
frmMain.MenuCreateNewTask.enabled := false;
frmMain.MenuFloatingVariables.enabled := false;
frmMain.MenuDocumentation.enabled := false;
frmMain.LabelFrontPageBuild.Caption := 'Version Build : ' + GetAppVersionStr;
frmMain.LabelScriptActiveRow.Caption := '';
frmMain.visible := true;
if frmMain.CBWelcomeStartup.Checked then TimerCounter := 31 else TimerCounter := 3;
  if LowerCase(ParamStr(1)) = 'nowelcome' then
  begin
   TimerCounter := 1;
   LogItStamp('Application startup parameter #1 = ' + ParamStr(1),2);
  end;
Timeout := 15;
MaxTimeout := 0;
ChillTime := 5;
verbose := false;
StringsToRemove := TStringList.Create;
if fileexists(extractfilepath(application.ExeName)+ '\Version History.txt') then frmMain.MemoVersionHistory.Lines.LoadFromFile(extractfilepath(application.ExeName)+ '\Version History.txt');
MenuAllreadyAdded := false;
NumberOfCommands := '74 script commands available - 21.06.2023';
end;

procedure TFrmMain.FormDestroy(Sender: TObject);
var
  i: Integer;
  sshClient: TScSSHClient;
  sshShell: TScSSHShell;
begin
{  for i := SSHClientList.Count - 1 downto 0 do
  begin
    sshClient := SSHClientList[i];
    sshShell := SSHShellList[i];
    if sshClient.Connected then
      sshClient.Disconnect;
    sshShell.Free;
    SSHShellList.Delete(i);
    sshClient.Free;
    SSHClientList.Delete(i);
  end;
  SSHShellList.Free;
  SSHClientList.Free; }
end;

end.
