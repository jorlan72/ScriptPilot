unit snippets;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Menus, Vcl.ComCtrls,
  Vcl.StdCtrls, FlexCel.Core, FlexCel.XlsAdapter, ShellAPI, System.UITypes, System.TimeSpan,
  System.Diagnostics, vcl.Clipbrd, System.RegularExpressions, vcl.ValEdit, comobj, IdHTTP, winsock,
  System.IOUtils, registry, System.Net.HttpClient, System.Net.URLClient, system.Win.Crtl, vcl.Grids,
  ScSSHChannel, ScSSHClient, ScBridge;

type
  TExternalProcess = record
   ProcessHandle: THandle;
   ExitCode: DWORD;
 end;

var
  TerminateExternalProcessFlag: Boolean = False;

procedure PopulateComboBoxWithSubfolders(const AFolder: string; ComboBox: TComboBox);
procedure PopulateComboBoxWithFilesOfType(const AFolder, AFileType: string; ComboBox: TComboBox);
Function GetAppVersionStr : string;
procedure LogItStamp(logtext : string; space : integer);
procedure LogIt(logtext : string; space : integer);
procedure ShellOpen(const Url: string; const Params: string = '');
Procedure InitializeVariablesAndData;
Procedure UpdateVariables;
procedure UpdateSettings;
Procedure WaitForContinue;
function FindExecutable(const AFileName: string): string;
function ShellExecuteAndWait(const AFileName: string; const AParameters: string = ''): TExternalProcess;
procedure TerminateExternalProcess(Process: TExternalProcess);
Procedure Chill(seconds : integer);
function SheetExists(const FileName, SheetName: string): Boolean;
function WaitForFile(const Directory, FileName: string; TimeoutSeconds: Integer): Boolean;
procedure OpenNotepadWithMemoText(Memo: TMemo);
function FetchTextAfterSearch(Memo: TMemo; FromLine, ToLine: Integer; const SearchTerm: string; FetchLength: Integer): string;
function TrimOccurrences(StringsToRemove: TStringList; const SourceText: string): string;
function RemoveAnsiEscapeCodes(const SourceText: string): string;
function RemoveNonPrintableChars(const SourceText: string): string;
function RemoveBlankLines(const SourceText: string): string;
procedure CleanMemoText(Memo: TMemo; StringsToRemove: TStringList);
function FindValueListEditorKey(ValueListEditor: TValueListEditor; const Key: string): Integer;
function GetPublicIPAddress: string;
function IsValidIPv4Address(const Address: string): Boolean;
function GetHostIP(const Hostname: string): string;
procedure SearchAndReplaceWord(const SourceFileName, DestinationFileName: String);
function CopyFileToTemp(const AFileName: string): string;
function DeleteFile(const AFileName: string): Boolean;
function CopyDirectory(const ASourceDir, ADestDir: string): Boolean;
function FindIndexByText(AComboBox: TComboBox; AText: string): Integer;
function CopyFileWithNewName(const ASourceFile, ADestDir, ANewName: string): Boolean;
function CreateDirectoryIfNotExists(const ADirectory: string): Boolean;
function ReadRegistryString(const RootKey: HKEY; const Key, ValueName: string; const DefaultValue: string = ''): string;
function WriteRegistryString(const RootKey: HKEY; const Key, ValueName, Value: string): Boolean;
procedure AddMenuItem(const Caption: string; const OnClick: TNotifyEvent);
procedure ConnectToTelnetHost(const Host: string);
Procedure DisconnectTelnetHost;
function IsWebsiteUp(const URL: string): Boolean;
function ExecuteCommand(const Command: string; OutputMemo: TMemo; OutputToMemo: Boolean): Boolean;
function ExecutePowerShellCommand(const Command: string; OutputMemo: TMemo; OutputToMemo: Boolean): Boolean;
procedure CopyTextFromWindowToMemo(const WindowToFind: string; const MemoToOutputText: TMemo);
procedure OpenNotepadWithValueListEditor(ValueListEditor: TValueListEditor);
procedure ModifyValueListEditorItem(editor: TValueListEditor; const keyName: string; const action: string; const newValue: string = '');
procedure LoadVariablesFromFile(const FileName: string);
Procedure SSHConnectEx(HostName : String; Username : String; Password : String);
Procedure SSHDisconnectEx(HostName: string);
Procedure SSHCommandEx(HostName : String; Command : String);

implementation

{$POINTERMATH ON}

 uses
   main;

function FindExecutable(const AFileName: string): string;
var
  Buffer: array [0 .. MAX_PATH - 1] of Char;
  FilePart: PChar;
begin
  if SearchPath(nil, PChar(AFileName), '.exe', MAX_PATH, Buffer, FilePart) <> 0 then
    Result := Buffer
  else
    Result := '';
end;

function ShellExecuteAndWait(const AFileName: string; const AParameters: string = ''): TExternalProcess;
var
  StartInfo: TStartupInfo;
  ProcessInfo: TProcessInformation;
  FullPath, CommandLine: string;
  WaitResult: DWORD;
begin
  FullPath := FindExecutable(AFileName);
  if FullPath = '' then
    raise Exception.CreateFmt('Unable to find executable "%s"', [AFileName]);
  CommandLine := FullPath + ' ' + AParameters;
  FillChar(StartInfo, SizeOf(TStartupInfo), #0);
  StartInfo.cb := SizeOf(TStartupInfo);
  StartInfo.dwFlags := STARTF_USESHOWWINDOW;
  StartInfo.wShowWindow := SW_SHOW;
  FillChar(ProcessInfo, SizeOf(TProcessInformation), #0);
  if CreateProcess(
      nil, PChar(CommandLine), nil, nil, False,
      CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS,
      nil, nil, StartInfo, ProcessInfo) then
  begin
    Result.ProcessHandle := ProcessInfo.hProcess;
    repeat
      WaitResult := MsgWaitForMultipleObjects(1, ProcessInfo.hProcess, False, INFINITE, QS_ALLINPUT);
      if WaitResult = WAIT_OBJECT_0 + 1 then
        Application.ProcessMessages;
      if TerminateExternalProcessFlag then
      begin
        TerminateProcess(ProcessInfo.hProcess, 0);
        Break;
      end;
    until WaitResult = WAIT_OBJECT_0;
    GetExitCodeProcess(ProcessInfo.hProcess, Result.ExitCode);
    CloseHandle(ProcessInfo.hThread);
    CloseHandle(ProcessInfo.hProcess);
  end
  else
  begin
    raise Exception.CreateFmt('Error executing "%s": %s', [FullPath, SysErrorMessage(GetLastError)]);
  end;
end;

procedure TerminateExternalProcess(Process: TExternalProcess);
begin
  TerminateProcess(Process.ProcessHandle, Process.ExitCode);
end;

procedure ShellOpen(const Url: string; const Params: string = '');
begin
  ShellExecute(0, 'open', PChar(Url), PChar(Params), nil, SW_SHOWNORMAL);
end;

procedure PopulateComboBoxWithSubfolders(const AFolder: string; ComboBox: TComboBox);
var
  SearchRec: TSearchRec;
  SubfolderName: string;
begin
  ComboBox.Clear; // clear existing items in ComboBox
  // search for subfolders in specified folder
  if FindFirst(AFolder + '\*', faDirectory, SearchRec) = 0 then
  begin
    repeat
      // skip '.' and '..' folders
      if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
      begin
        // check if current item is a directory (subfolder)
        if (SearchRec.Attr and faDirectory) = faDirectory then
        begin
          SubfolderName := SearchRec.Name;
          // add subfolder name to ComboBox
          ComboBox.Items.Add(SubfolderName);
        end;
      end;
    until FindNext(SearchRec) <> 0;
    FindClose(SearchRec);
  end;
 if ComboBox.Items.Count > 0 then ComboBox.ItemIndex := 0;
end;

procedure PopulateComboBoxWithFilesOfType(const AFolder, AFileType: string; ComboBox: TComboBox);
var
  SearchRec: TSearchRec;
  FileName: string;
  ShortName : string;
begin
  // clear existing items in ComboBox
  ComboBox.Clear;
  // search for files of specified type in specified folder
  if FindFirst(AFolder + '\' + AFileType, faAnyFile, SearchRec) = 0 then
  begin
    repeat
      // add file name to ComboBox
      FileName := SearchRec.Name;
      ShortName := StringReplace(FileName, '.xlsx','', [rfReplaceAll]);
      ComboBox.Items.Add(ShortName);
    until FindNext(SearchRec) <> 0;
    FindClose(SearchRec);
  end;
 if ComboBox.Items.Count > 0 then ComboBox.ItemIndex := 0;
end;

function GetAppVersionStr: string;
var
  Exe: string;
  Size, Handle: DWORD;
  Buffer: TBytes;
  FixedPtr: PVSFixedFileInfo;
begin
  Exe := ParamStr(0);
  Size := GetFileVersionInfoSize(PChar(Exe), Handle);
  if Size = 0 then
    RaiseLastOSError;
  SetLength(Buffer, Size);
  if not GetFileVersionInfo(PChar(Exe), Handle, Size, Buffer) then
    RaiseLastOSError;
  if not VerQueryValue(Buffer, '\', Pointer(FixedPtr), Size) then
    RaiseLastOSError;
  Result := Format('%d.%d.%d.%d',
    [LongRec(FixedPtr.dwFileVersionMS).Hi,  //major
     LongRec(FixedPtr.dwFileVersionMS).Lo,  //minor
     LongRec(FixedPtr.dwFileVersionLS).Hi,  //release
     LongRec(FixedPtr.dwFileVersionLS).Lo]) //build
end;

procedure LogItStamp(logtext : string; space : integer);
var
 counter : integer;
begin
if not WillLog then exit;
  frmMain.MemoLog.Lines.Add(datetimetostr(now) + ' : ' + logtext);
    if space > 0 then begin
      for counter := 1 to space do begin
      frmMain.MemoLog.Lines.Add('');
      end;
    end;
end;

procedure LogIt(logtext : string; space : integer);
var
 counter : integer;
begin
if not WillLog then exit;
  frmMain.MemoLog.Lines.Add(logtext);
    if space > 0 then begin
      for counter := 1 to space do begin
      frmMain.MemoLog.Lines.Add('');
      end;
    end;
end;

Procedure InitializeVariablesAndData;
var
 xls : TXlsFile;
 Counter : integer;
 FileName : string;
begin
xls := TXlsFile.Create;
   if newproject then
      begin
       frmMain.DropDownProject.ItemIndex := FindIndexByText(frmMain.DropDownProject, NewProjectText);
      end;
filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + 'data.xlsx';
CopyFileToTemp(filename);
filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + 'data.tmp';
  if SheetExists(filename,'Data') then begin
   xls.open(filename);
   xls.ActiveSheetByName := 'Data';
   LogItStamp('Activating project "' + frmmain.DropDownProject.Text + '"',0);
   LogItStamp('To open the "projects" folder, hit F2',0);
   LogItStamp('Adding source data to drop-down from row 11 in the Excel "Data" sheet in the project "Data.xlsx" file',0);
    Counter := 11;
      while not xls.GetCellValue(counter,1).IsEmpty do begin
       frmMain.DropDownData.Items.add(xls.GetCellValue(Counter,1));
       inc(Counter);
      end;
       if frmMain.DropDownData.items.Count > 0 then frmMain.DropDownData.ItemIndex := 0;
       xls.free;
       // Add global variables
         frmMain.ValueListEditorVariables.Strings.Clear;
         frmmain.ValueListEditorVariables.InsertRow('#DataSourceSelected#', frmMain.DropDownData.text, true);
         frmmain.ValueListEditorVariables.InsertRow('#DataSourceIndex#', inttostr(frmMain.DropDownData.ItemIndex + 1), true);
         frmmain.ValueListEditorVariables.InsertRow('#DataSourceItemsCount#', inttostr(frmmain.DropDownData.Items.Count), true);
         frmmain.ValueListEditorVariables.InsertRow('#ApplicationPath#', extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text, true);
         UpdateVariables;
  end else begin
   LogItStamp('Activating project "' + frmmain.DropDownProject.Text + '"',0);
   LogItStamp('ERROR: "Data.xlsx" or sheet "Data" missing. Skipping updating project',0);
   LogItStamp('If files have been manually deleted in the "Projects" folder, while the application is running, try a refresh from the Edit menu (or hit F5)',0);
   frmMain.MenuProjectDocuments.Clear;
  end;
DeleteFile(filename);
end;

Procedure UpdateVariables;
var
 xls : TXlsFile;
 Counter : integer;
 FileName : string;
 FetchVariable : variant;
 FetchValue : variant;
 KeyIndex : Integer;
begin
xls := TXlsFile.Create;
filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + 'data.xlsx';
CopyFileToTemp(filename);
filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + 'data.tmp';
  if SheetExists(filename,'Data') then begin
    xls.open(filename);
    xls.ActiveSheetByName := 'Data';
    LogItStamp('Initializing variables from selection - "' + frmMain.DropDownData.text + '"',0);
    LogItStamp('To edit the data, hit F3 (Microsoft Excel required)',0);
    Counter := 2;
       while not xls.GetCellValue(10, Counter).IsEmpty do
       begin
         FetchVariable := xls.GetCellValue(10, Counter);
         FetchValue := xls.GetCellValue(frmMain.DropDownData.ItemIndex + 11, Counter);
         keyindex := FindValueListEditorKey(frmMain.ValueListEditorVariables, FetchVariable);
          if keyindex = -1 then frmmain.ValueListEditorVariables.InsertRow(FetchVariable, FetchValue, true)
                           else frmMain.ValueListEditorVariables.Values[FetchVariable] := FetchValue;
         inc(Counter);
       end;
  frmMain.ValueListEditorVariables.Values['#DataSourceSelected#'] := frmMain.DropDownData.Text;
  frmmain.ValueListEditorVariables.Values['#DataSourceIndex#'] := inttostr(frmMain.DropDownData.ItemIndex + 1);
  end else begin
   LogItStamp('ERROR: "Data.xlsx" or sheet "Data" missing. Skipping updating variables',0);
   LogItStamp('If files have been manually deleted in the "Projects" folder, while the application is running, try a refresh from the Edit menu (or hit F5)',0);
   frmMain.MenuProjectDocuments.Clear;
  end;
   if not MenuAllreadyAdded then
    begin
      frmMain.MenuProjectDocuments.Clear;
     if SheetExists(filename,'Project documents') then
       begin
         xls.ActiveSheetByName := 'Project documents';
         LogItStamp('Adding "Project documents" to the menu. You can edit the content in the "Data.xlsx" file, in the "Project documents" sheet.',0);
         Counter := 6;
           while not xls.GetCellValue(Counter, 1).IsEmpty do
            begin
            LogItStamp('Adding menu item : "' + vartostr(xls.GetCellValue(Counter, 1)) + '"',0);
            MenuTag := vartostr(xls.GetCellValue(Counter, 2));
            AddMenuItem(vartostr(xls.GetCellValue(Counter, 1)), frmMain.MenuItemClick);
            inc(Counter);
            end;
         MenuAllreadyAdded := true;
       end else
       begin
        LogItStamp('No Project documents found',0);
        MenuAllreadyAdded := true;
       end;
    end;
 xls.free;
 DeleteFile(filename);
end;

procedure UpdateSettings;
var
 xls : TXlsFile;
 FileName : string;
 TempString : String;
 row, inputlines : Integer;
begin
frmmain.ValueListEditorInput.Strings.Clear;
TimeOut := 30;
MaxTimeout := 0;
xls := TXlsFile.Create;
filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.xlsx';
CopyFileToTemp(filename);
filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.tmp';
  if SheetExists(filename,'Settings') then begin
    xls.open(filename);
    xls.ActiveSheetByName := 'Settings';
    LogItStamp('Reading settings for task "' + frmMain.DropDownTask.text + '"',0);
    LogItStamp('To edit task, hit F4 (Microsoft Excel required)',0);
      if not xls.getcellvalue(3,2).IsEmpty then begin
       TempString := vartostr(xls.getcellvalue(3,2));
       frmMain.LabelTaskInformation.Caption := TempString;
      end;
      if not xls.getcellvalue(6,2).IsEmpty then begin
       TempString := vartostr(xls.getcellvalue(6,2));
       if not TryStrtoInt(Tempstring, Timeout) then Timeout := 5;
      end;
      if not xls.getcellvalue(7,2).IsEmpty then begin
       TempString := vartostr(xls.getcellvalue(7,2));
       if not TryStrtoInt(TempString, MaxTimeout) then MaxTimeout := 0;
      end;
      if not xls.getcellvalue(8,2).IsEmpty then begin
       TempString := vartostr(xls.getcellvalue(8,2));
       if not trystrtoint(TempString, ChillTime) then ChillTime := 4;
      end;
        row := 20;
        inputlines := 0;
        SetLength(InputArray, inputlines);
           while not xls.GetCellValue(row, 1).IsEmpty do begin
             inc(inputlines);
             SetLength(InputArray, inputlines);
             InputArray[inputlines - 1].InputDescription := xls.GetCellValue(row, 1);
             InputArray[inputlines - 1].VariableName := xls.GetCellValue(row, 2);
             InputArray[inputlines - 1].InitialValue := xls.GetCellValue(row, 3);
             frmmain.ValueListEditorInput.InsertRow(InputArray[inputlines - 1].InputDescription, InputArray[inputlines - 1].InitialValue, true);
             inc(row);
           end;
   end else begin
             if not frmMain.DropDownTask.ItemIndex = -1 then
              LogItStamp('ERROR: Task file missing or sheet "Settings" missing. Skipping updating settings',0);
              LogItStamp('If files have been manually deleted in the "Projects" folder, while the application is running, try a refresh from the Edit menu (or hit F5)',0);
            end;
   if SheetExists(filename,'Trim') then
   begin
        xls.open(filename);
        xls.ActiveSheetByName := 'Trim';
        StringsToRemove.Clear;
        row := 5;
        TempString := '';
             while not xls.GetCellValue(row, 1).IsEmpty do begin
              TempString := vartostr(xls.GetCellValue(row, 1));
              StringsToRemove.Add(tempString);
              inc(row);
             end;
   end;
 xls.free;
 DeleteFile(filename);
end;

Procedure WaitForContinue;
begin
 ContinuePressed := false;
   while not ContinuePressed do begin
     sleep(100);
     application.ProcessMessages;
     if stoppressed then exit;
   end;
if stoppressed then exit;
end;

Procedure Chill(seconds : integer);
var
 interval : integer;
 counter : integer;
begin
  interval := seconds * 10;
  for counter := 0 to interval do begin
  sleep(100);
  application.ProcessMessages;
  if stoppressed then exit;
  end;
end;

function SheetExists(const FileName, SheetName: string): Boolean;
var
  ExcelFile: TExcelFile;
  SheetIndex: Integer;
  FileOpened: Boolean;
begin
  Result := False;
  // Check if the file exists
  if not FileExists(FileName) then
    Exit;
  // Create an instance of TExcelFile
  ExcelFile := TXlsFile.Create;
  try
    FileOpened := False;
    while not FileOpened do
    begin
      // Try to open the Excel file
      try
        ExcelFile.Open(FileName);
        FileOpened := True;
      except
        on E: Exception do
        begin
          // If the file is locked, show a message and wait for the OK button
          if E is EStreamError then
          begin
            if MessageDlg('The Excel file needed is currently open. Please close it and click OK to continue.',
              mtInformation, [mbOK, mbCancel], 0) = mrCancel then
            begin
              // StopPressed := True;
              LogItStamp('ERROR: The Excel file needed is currently open. User selected the Cancel button',0);
              Exit;
            end;
          end
          else
            raise;
        end;
      end;
    end;
    // Try to find the sheet by its name
    SheetIndex := ExcelFile.GetSheetIndex(SheetName, false);
    // Check if the sheet exists (SheetIndex will be >= 1)
    Result := SheetIndex >= 1;
  finally
    // Free the TExcelFile instance
    ExcelFile.Free;
  end;
end;

function WaitForFile(const Directory, FileName: string; TimeoutSeconds: Integer): Boolean;
var
  FilePath: string;
  Stopwatch: TStopwatch;
  TimeoutEnabled: Boolean;
begin
  Result := False;
  FilePath := IncludeTrailingPathDelimiter(Directory) + FileName;
  // Start the stopwatch to keep track of the elapsed time
  Stopwatch := TStopwatch.StartNew;
  TimeoutEnabled := TimeoutSeconds > 0;
  // Keep checking for the file until the timeout is reached or indefinitely if TimeoutSeconds is 0
  while (not TimeoutEnabled) or (Stopwatch.Elapsed.TotalSeconds < TimeoutSeconds) do
  begin
    // If the file is found, set the result to True and exit the loop
    if FileExists(FilePath) then
    begin
      Result := True;
      Break;
    end;
    // Sleep for a short period (e.g., 100 milliseconds) to avoid high CPU usage
    Sleep(100);
    if StopPressed then exit;
    Application.Processmessages;
  end;
end;

procedure OpenNotepadWithMemoText(Memo: TMemo);
var
  NotepadHandle, EditHandle: HWND;
begin
  // Launch a new instance of Notepad
  if ShellExecute(0, 'open', 'notepad.exe', nil, nil, SW_SHOWNORMAL) <= 32 then
  begin
    raise Exception.Create('Unable to open Notepad.');
    Exit;
  end;
  // Wait for the new Notepad instance to be launched
  repeat
    Sleep(100);
    NotepadHandle := FindWindow('Notepad', nil);
  until NotepadHandle <> 0;
  // Bring the new Notepad instance to the foreground
  SetForegroundWindow(NotepadHandle);
  // Attempt to get the handle of the Edit control for Windows 11
  EditHandle := FindWindowEx(NotepadHandle, 0, 'RichEdit', nil);
  // Fallback to the Windows 10 class name if necessary
  if EditHandle = 0 then
    EditHandle := FindWindowEx(NotepadHandle, 0, 'Edit', nil);
  if EditHandle = 0 then
  begin
    ShowMessage('Windows 11 has a new rich edit Notepad that is not supported by ScriptPilot.');
    Exit;
  end;
  // Copy the text from the Memo to the clipboard
  Clipboard.AsText := Memo.Lines.Text;
  // Set focus to the Edit control in Notepad
  SetFocus(EditHandle);
  // Paste the text from the clipboard into the Edit control
  SendMessage(EditHandle, WM_PASTE, 0, 0);
  // Move the cursor to the beginning of the Edit control
  SendMessage(EditHandle, EM_SETSEL, 0, 0);
  // Scroll the document to the top
  SendMessage(EditHandle, EM_SCROLLCARET, 0, 0);
end;

function FetchTextAfterSearch(Memo: TMemo; FromLine, ToLine: Integer; const SearchTerm: string; FetchLength: Integer): string;
var
  LineIndex, SearchPos: Integer;
  LineText: string;
begin
  Result := '';
  // Ensure FromLine and ToLine are within the valid range
  if (FromLine < 0) or (ToLine >= Memo.Lines.Count) or (FromLine > ToLine) then
  begin
    raise Exception.Create('Invalid FromLine or ToLine.');
    Exit;
  end;
  // Iterate through the specified lines
  for LineIndex := FromLine to ToLine do
  begin
    LineText := Memo.Lines[LineIndex];
    SearchPos := Pos(SearchTerm, LineText);
    // If the search term is found in the line
    if SearchPos > 0 then
    begin
      // Fetch the specified number of characters after the search term
      Result := Copy(LineText, SearchPos + Length(SearchTerm), FetchLength);
      Break;
    end;
  end;
end;

function TrimOccurrences(StringsToRemove: TStringList; const SourceText: string): string;
var
  Index: Integer;
  TempText: string;
  RegEx: TRegEx;
begin
  TempText := SourceText;
  // Remove specified strings
  for Index := 0 to StringsToRemove.Count - 1 do
  begin
    TempText := StringReplace(TempText, StringsToRemove[Index], '', [rfReplaceAll]);
  end;
  // Replace consecutive line breaks with a single one
  RegEx := TRegEx.Create('(\r\n|\r|\n){2,}');
  TempText := RegEx.Replace(TempText, sLineBreak);
  Result := TempText;
end;

function RemoveAnsiEscapeCodes(const SourceText: string): string;
var
  RegEx: TRegEx;
begin
  // Remove common ANSI escape sequences
  RegEx := TRegEx.Create('(\x1B\[|\x1B\()(?:[0-9]{1,2}(?:;[0-9]{1,2})?)?[m|K]?');
  Result := RegEx.Replace(SourceText, '');

  // Remove ESC character (ASCII 27) followed by a left square bracket ([)
  Result := StringReplace(Result, #27 + '[', '', [rfReplaceAll]);
end;

function RemoveNonPrintableChars(const SourceText: string): string;
var
  RegEx: TRegEx;
begin
  // Remove ANSI escape sequences
  RegEx := TRegEx.Create('\x1B\[[^A-Za-z]*[A-Za-z]');
  Result := RegEx.Replace(SourceText, '');

  // Remove non-breaking space (chr(160)), zero-width space (chr(8203)), and tab (chr(9))
  Result := StringReplace(Result, Chr(160), '', [rfReplaceAll]);
  Result := StringReplace(Result, Chr(8203), '', [rfReplaceAll]);
  Result := StringReplace(Result, Chr(9), '', [rfReplaceAll]);

  // Replace multiple line breaks with a single line break (first pass)
  RegEx := TRegEx.Create('(\r\n){2,}');
  Result := RegEx.Replace(Result, sLineBreak);

  // Replace multiple line breaks with a single line break (second pass)
  RegEx := TRegEx.Create('(\r\n){2,}');
  Result := RegEx.Replace(Result, sLineBreak);
end;

function RemoveBlankLines(const SourceText: string): string;
var
  Source, Dest: TStringList;
  i: Integer;
begin
  Source := TStringList.Create;
  Dest := TStringList.Create;
  try
    Source.Text := SourceText;
    for i := 0 to Source.Count - 1 do
    begin
      if Trim(Source[i]) <> '' then
        Dest.Add(Source[i]);
    end;
    Result := Dest.Text;
  finally
    Source.Free;
    Dest.Free;
  end;
end;

procedure CleanMemoText(Memo: TMemo; StringsToRemove: TStringList);
var
  CleanedText: string;
begin
  CleanedText := Memo.Text;
  CleanedText := TrimOccurrences(StringsToRemove, CleanedText);
  CleanedText := RemoveAnsiEscapeCodes(CleanedText);
  CleanedText := RemoveNonPrintableChars(CleanedText);
  CleanedText := RemoveBlankLines(CleanedText);
  Memo.Text := CleanedText;
end;

function FindValueListEditorKey(ValueListEditor: TValueListEditor; const Key: string): Integer;
var
  I: Integer;
begin
result := -1;
  for I := 0 to ValueListEditor.Strings.Count - 1 do
  begin
    if ValueListEditor.Strings.KeyNames[I] = Key then
    begin
      Result := I;
      exit
        end else Result := -1;
  end;
end;

function GetPublicIPAddress: string;
var
  IdHTTP: TIdHTTP;
  Resp: TStringStream;
begin
  Result := '';
  try
    IdHTTP := TIdHTTP.Create;
    Resp := TStringStream.Create('');
    try
      IdHTTP.Get('http://api.ipify.org', Resp);
      Result := Trim(Resp.DataString);
    finally
      Resp.Free;
      IdHTTP.Free;
    end;
  except
    on E: Exception do
    begin
      LogItStamp('ERROR: not connected to the Internet',0);
      Result := '';
    end;
  end;
end;

function IsValidIPv4Address(const Address: string): Boolean;
var
  i, Num, Count: Integer;
  NumStr, Temp: string;
begin
  Result := False;
  Count := 0;
  Temp := Address + '.';
  for i := 1 to Length(Temp) do
  begin
    if Temp[i] = '.' then
    begin
      Inc(Count);
      if Count > 4 then Exit;
      Num := StrToIntDef(NumStr, -1);
      if (Num < 0) or (Num > 255) then Exit;
      NumStr := '';
    end
    else if not CharInSet(Temp[i], ['0'..'9']) then Exit
    else NumStr := NumStr + Temp[i];
  end;
  Result := Count = 4;
end;

procedure SearchAndReplaceWord(const SourceFileName, DestinationFileName: String);
var
  Counter: Integer;
  TargetText, ReplacementText: String;
  WordApplication, WordFile: Variant;
  FoundAndReplaced : boolean;
begin
  if not FileExists(SourceFileName) then begin
    LogItStamp('ERROR: The source word file does not exist. Row: ' + inttostr(currentrow),0);
    Exit;
  end;

  try
    WordApplication := CreateOleObject('Word.Application');
    if not VarIsNull(WordApplication) then begin
      WordApplication.Visible := True;
      WordApplication.DisplayAlerts := False;
      WordFile := WordApplication.Documents.Open(SourceFileName);
      if not VarIsNull(WordFile) then begin
        for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count - 1 do begin
          TargetText := frmMain.ValueListEditorVariables.Strings.KeyNames[Counter];
          ReplacementText := frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter];
           WordFile.Content.Find.ClearFormatting;
           WordFile.Content.Find.Replacement.ClearFormatting;
           FoundAndReplaced := False;
           FoundAndReplaced := WordFile.Content.Find.Execute(TargetText,False,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,ReplacementText,2,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
          Application.ProcessMessages;
        end;
      end;

      if not DirectoryExists(ExtractFilePath(DestinationFileName)) then begin
         LogItStamp('ERROR: The destination folder does not exist. Row: ' + inttostr(currentrow),0);
        Exit;
      end;

      WordFile.SaveAs(DestinationFileName);
      LogItStamp('Word file string replaced and saved to ' + DestinationFileName,0);
      WordFile.Close;
      WordApplication.Quit;
      WordFile := Unassigned;
    end;
  finally
    WordApplication := Unassigned;
  end;
end;

function GetHostIP(const Hostname: string): string;
var
  WSAData: TWSAData;
  HostEnt: PHostEnt;
  Addr: PAnsiChar;
  i: Integer;
begin
  Result := '';
  if WSAStartup(MAKEWORD(2, 2), WSAData) <> 0 then Exit;
  try
    HostEnt := gethostbyname(PAnsiChar(AnsiString(Hostname)));
    if HostEnt <> nil then
    begin
      i := 0;
      while HostEnt^.h_addr_list[i] <> nil do
      begin
        Addr := HostEnt^.h_addr_list[i];
        Result := Format('%d.%d.%d.%d', [Ord(Addr[0]), Ord(Addr[1]), Ord(Addr[2]), Ord(Addr[3])]);
        Inc(i);
      end;
    end;
  finally
    WSACleanup;
  end;
end;

function CopyFileToTemp(const AFileName: string): string;
var
  SourceFileName, DestFileName: string;
begin
  SourceFileName := AFileName;
  DestFileName := ChangeFileExt(SourceFileName, '.tmp');
  Result := ExtractFilePath(SourceFileName) + ExtractFileName(DestFileName);
  try
    if TFile.Exists(SourceFileName) then // <-- check if source file exists before attempting to copy
      TFile.Copy(SourceFileName, Result, True) // <-- added "True" parameter to overwrite existing file
    else
      Result := ''; // <-- set empty string if file doesn't exist
  except
    on E: Exception do
      ShowMessage('Error copying file: ' + E.Message);
  end;
end;

function DeleteFile(const AFileName: string): Boolean;
begin
  Result := False;
  try
    if TFile.Exists(AFileName) then // <-- check if file exists before attempting to delete
    begin
      TFile.Delete(AFileName);
      Result := True;
    end
    else
      Result := True; // <-- set to True if file doesn't exist
  except
    on E: Exception do
      ShowMessage('Error deleting file: ' + E.Message);
  end;
end;

function CopyDirectory(const ASourceDir, ADestDir: string): Boolean;
var
  SourceDir, DestDir: string;
  FileInfo: TSearchRec;
  FindResult: Integer;
  NewDestDir: string;
begin
  SourceDir := IncludeTrailingPathDelimiter(ASourceDir);
  DestDir := IncludeTrailingPathDelimiter(ADestDir);
  Result := False;
  try
    // Create the destination directory
    if not CreateDir(DestDir) then
      Exit;
    // Find the first file/directory in the source directory
    FindResult := FindFirst(SourceDir + '*.*', faAnyFile, FileInfo);
    while FindResult = 0 do
    begin
      if (FileInfo.Name <> '.') and (FileInfo.Name <> '..') then
      begin
        if (FileInfo.Attr and faDirectory) = faDirectory then
        begin
          // Recursively copy subdirectory and its contents
          NewDestDir := IncludeTrailingPathDelimiter(DestDir + FileInfo.Name);
          if not CopyDirectory(SourceDir + FileInfo.Name, NewDestDir) then
            Exit;
        end
        else
        begin
          // Copy file to destination directory
          if not CopyFile(PChar(SourceDir + FileInfo.Name), PChar(DestDir + FileInfo.Name), False) then
            Exit;
        end;
      end;
      FindResult := FindNext(FileInfo);
    end;
    FindClose(FileInfo);
    Result := True;
  except
    on E: Exception do
      LogItStamp('ERROR: Creating new project failed :' + E.Message,0);
  end;
end;

function FindIndexByText(AComboBox: TComboBox; AText: string): Integer;
begin
  for Result := 0 to AComboBox.Items.Count - 1 do
  begin
    if SameText(AComboBox.Items[Result], AText) then
      Exit;
  end;
  Result := -1; // not found
end;

function CopyFileWithNewName(const ASourceFile, ADestDir, ANewName: string): Boolean;
var
  DestFile: string;
begin
  Result := False;
  try
    if not DirectoryExists(ADestDir) then
      Exit;
    DestFile := IncludeTrailingPathDelimiter(ADestDir) + ANewName;
    if FileExists(DestFile) then
      Exit;
    if not CopyFile(PChar(ASourceFile), PChar(DestFile), False) then
      Exit;
    Result := True;
  except
    on E: Exception do
      LogItStamp('ERROR: Copying file failed :' + E.Message,0);
  end;
end;

function CreateDirectoryIfNotExists(const ADirectory: string): Boolean;
begin
  Result := DirectoryExists(ADirectory);
  if not Result then
  begin
    try
      CreateDir(ADirectory);
      Result := True;
    except
      on E: Exception do
        LogItStamp('ERROR: Creating directory :' + E.Message,0);
    end;
  end
  else
  begin
    // Directory already exists
    Result := False;
  end;
end;

function ReadRegistryString(const RootKey: HKEY; const Key, ValueName: string; const DefaultValue: string = ''): string;
var
  Reg: TRegistry;
begin
  Result := DefaultValue;
  Reg := TRegistry.Create;
  try
    Reg.RootKey := RootKey;
    if Reg.OpenKeyReadOnly(Key) then
    begin
      Result := Reg.ReadString(ValueName);
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

function WriteRegistryString(const RootKey: HKEY; const Key, ValueName, Value: string): Boolean;
var
  Reg: TRegistry;
begin
  Result := False;
  Reg := TRegistry.Create;
  try
    Reg.RootKey := RootKey;
    if Reg.OpenKey(Key, True) then
    begin
      Reg.WriteString(ValueName, Value);
      Reg.CloseKey;
      Result := True;
    end;
  finally
    Reg.Free;
  end;
end;

procedure AddMenuItem(const Caption: string; const OnClick: TNotifyEvent);
var
  MenuItem: TMenuItem;
begin
  MenuItem := TMenuItem.Create(frmMain.MenuProjectDocuments);
  MenuItem.Caption := Caption;
  MenuItem.Hint := MenuTag;
  MenuItem.OnClick := OnClick;
  frmMain.MenuProjectDocuments.Add(MenuItem);
end;

procedure ConnectToTelnetHost(const Host: string);
begin
 if frmMain.TelnetClient01.isConnected then
  begin
   LogItStamp('Telnet connection to another host already active. Ignoring command at row ' + inttostr(CurrentRow),0);
   Exit;
  end;
    LogItStamp('Trying Telnet connection to host ' + ExcelParam01 + ' with timeout of ' + inttostr(ChillTime) + ' seconds',0);
    frmMain.Telnetclient01.host := Host;
    frmMain.Telnetclient01.Port := '23';
    frmMain.Telnetclient01.Connect;
    Chill(ChillTime);
       if not frmMain.TelnetClient01.IsConnected then
        begin
         LogItStamp('ERROR: Could not connect to Telnet host ' + VarToStr(ExcelParam01), 0);
         exit
        end;
LogItStamp('Connected to Telnet host ' + VarToStr(ExcelParam01), 0);
end;

Procedure DisconnectTelnetHost;
begin
  frmMain.TelnetClient01.close;
  LogItStamp('Telnet host disconnected',0);
end;

function IsWebsiteUp(const URL: string): Boolean;
var
  HTTPClient: THTTPClient;
  HTTPResponse: IHTTPResponse;
begin
  HTTPClient := THTTPClient.Create;
  try
    try
      HTTPResponse := HTTPClient.Get(URL);
      Result := (HTTPResponse.StatusCode >= 200) and (HTTPResponse.StatusCode < 300);
    except
      on E: ENetHTTPClientException do
        Result := False;
    end;
  finally
    HTTPClient.Free;
  end;
end;

function ExecuteCommand(const Command: string; OutputMemo: TMemo; OutputToMemo: Boolean): Boolean;
var
  StartupInfo: TStartupInfo;
  ProcessInfo: TProcessInformation;
  SecurityAttr: TSecurityAttributes;
  hReadPipe, hWritePipe: THandle;
  dwBytesRead: DWORD;
  OutputBuffer: array [0 .. 4096] of AnsiChar;
  OutputString: AnsiString;
  Overlapped: TOverlapped;
begin
  Result := False;
  SecurityAttr.nLength := SizeOf(TSecurityAttributes);
  SecurityAttr.bInheritHandle := True;
  SecurityAttr.lpSecurityDescriptor := nil;
  if not CreatePipe(hReadPipe, hWritePipe, @SecurityAttr, 0) then
    raise Exception.Create('Failed to create pipe.');
  ZeroMemory(@StartupInfo, SizeOf(StartupInfo));
  StartupInfo.cb := SizeOf(StartupInfo);
  StartupInfo.dwFlags := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
  StartupInfo.wShowWindow := SW_HIDE;
  StartupInfo.hStdInput := GetStdHandle(STD_INPUT_HANDLE);
  StartupInfo.hStdOutput := hWritePipe;
  StartupInfo.hStdError := hWritePipe;
  ZeroMemory(@ProcessInfo, SizeOf(ProcessInfo));
  if not CreateProcess(nil, PChar('cmd.exe /c ' + Command),
    nil, nil, True, 0, nil, nil, StartupInfo, ProcessInfo) then
    raise Exception.Create('Failed to create process.');
  CloseHandle(hWritePipe);
  if OutputToMemo and Assigned(OutputMemo) then
  begin
   //
  end;
  FillChar(Overlapped, SizeOf(Overlapped), 0);
  Overlapped.hEvent := CreateEvent(nil, True, False, nil);
  while ReadFile(hReadPipe, OutputBuffer, SizeOf(OutputBuffer), dwBytesRead, @Overlapped) or (GetLastError = ERROR_IO_PENDING) do
  begin
    if GetLastError = ERROR_IO_PENDING then
    begin
      WaitForSingleObject(Overlapped.hEvent, INFINITE);
      GetOverlappedResult(hReadPipe, Overlapped, dwBytesRead, False);
    end;
    if dwBytesRead = 0 then
      Break;
    OutputString := AnsiString(Copy(OutputBuffer, 0, dwBytesRead));
    if OutputToMemo and Assigned(OutputMemo) then
    begin
      OutputMemo.Lines.Add(string(OutputString));
    end;
  end;
  CloseHandle(hReadPipe);
  CloseHandle(ProcessInfo.hThread);
  CloseHandle(ProcessInfo.hProcess);
  Result := True;
end;

function ExecutePowerShellCommand(const Command: string; OutputMemo: TMemo; OutputToMemo: Boolean): Boolean;
var
  StartupInfo: TStartupInfo;
  ProcessInfo: TProcessInformation;
  SecurityAttr: TSecurityAttributes;
  hReadPipe, hWritePipe: THandle;
  dwBytesRead: DWORD;
  OutputBuffer: array [0 .. 4096] of AnsiChar;
  OutputString: AnsiString;
  Overlapped: TOverlapped;
begin
  Result := False;
  SecurityAttr.nLength := SizeOf(TSecurityAttributes);
  SecurityAttr.bInheritHandle := True;
  SecurityAttr.lpSecurityDescriptor := nil;
  if not CreatePipe(hReadPipe, hWritePipe, @SecurityAttr, 0) then
    raise Exception.Create('Failed to create pipe.');
  ZeroMemory(@StartupInfo, SizeOf(StartupInfo));
  StartupInfo.cb := SizeOf(StartupInfo);
  StartupInfo.dwFlags := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
  StartupInfo.wShowWindow := SW_HIDE;
  StartupInfo.hStdInput := GetStdHandle(STD_INPUT_HANDLE);
  StartupInfo.hStdOutput := hWritePipe;
  StartupInfo.hStdError := hWritePipe;
  ZeroMemory(@ProcessInfo, SizeOf(ProcessInfo));
  if not CreateProcess(nil, PChar('powershell.exe -ExecutionPolicy Bypass -Command "' + Command + '"'),
    nil, nil, True, 0, nil, nil, StartupInfo, ProcessInfo) then
    raise Exception.Create('Failed to create process.');
  CloseHandle(hWritePipe);
  if OutputToMemo and Assigned(OutputMemo) then
  begin
   //
  end;
  FillChar(Overlapped, SizeOf(Overlapped), 0);
  Overlapped.hEvent := CreateEvent(nil, True, False, nil);
  while ReadFile(hReadPipe, OutputBuffer, SizeOf(OutputBuffer), dwBytesRead, @Overlapped) or (GetLastError = ERROR_IO_PENDING) do
  begin
    if GetLastError = ERROR_IO_PENDING then
    begin
      WaitForSingleObject(Overlapped.hEvent, INFINITE);
      GetOverlappedResult(hReadPipe, Overlapped, dwBytesRead, False);
    end;
    if dwBytesRead = 0 then
      Break;
    OutputString := AnsiString(Copy(OutputBuffer, 0, dwBytesRead));
    if OutputToMemo and Assigned(OutputMemo) then
    begin
     OutputMemo.Lines.Add(string(OutputString));
    end;
  end;
  CloseHandle(hReadPipe);
  CloseHandle(ProcessInfo.hThread);
  CloseHandle(ProcessInfo.hProcess);
  Result := True;
end;

procedure CopyTextFromWindowToMemo(const WindowToFind: string; const MemoToOutputText: TMemo);
var
  _hWnd, hWndMemo: HWND;
  buffer: array[0..255] of Char;
begin
  // Find the window
  _hWnd := FindWindow(nil, PChar(WindowToFind));
  if _hWnd = 0 then
  begin
    LogItStamp('ERROR: External Window not found.',0);
    Exit;
  end;

  // Find the memo in your application
  hWndMemo := MemoToOutputText.Handle;

  // Activate the window
  SetForegroundWindow(_hWnd);

  // Send Ctrl+A to select all text in the window
  keybd_event(VK_CONTROL, 0, 0, 0);
  keybd_event(Ord('A'), 0, 0, 0);
  keybd_event(Ord('A'), 0, KEYEVENTF_KEYUP, 0);
  keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);

  // Send Ctrl+C to copy the selected text in the window
  keybd_event(VK_CONTROL, 0, 0, 0);
  keybd_event(Ord('C'), 0, 0, 0);
  keybd_event(Ord('C'), 0, KEYEVENTF_KEYUP, 0);
  keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);

  // Wait for a short time for the copied text to be available
  Sleep(100);

  // Get the copied text from the clipboard
  if GetClipboardData(CF_TEXT) <> 0 then
  begin
    // Get the text from the clipboard and set it in the memo
    GetWindowText(hWndMemo, buffer, SizeOf(buffer));
    MemoToOutputText.Lines.Add(buffer);
  end;
end;

procedure OpenNotepadWithValueListEditor(ValueListEditor: TValueListEditor);
var
  NotepadHandle, EditHandle: HWND;
begin
  // Launch a new instance of Notepad
  if ShellExecute(0, 'open', 'notepad.exe', nil, nil, SW_SHOWNORMAL) <= 32 then
  begin
    raise Exception.Create('Unable to open Notepad.');
    Exit;
  end;
  // Wait for the new Notepad instance to be launched
  repeat
    Sleep(100);
    NotepadHandle := FindWindow('Notepad', nil);
  until NotepadHandle <> 0;
  // Bring the new Notepad instance to the foreground
  SetForegroundWindow(NotepadHandle);
  // Attempt to get the handle of the Edit control for Windows 11
  EditHandle := FindWindowEx(NotepadHandle, 0, 'RichEdit', nil);
  // Fallback to the Windows 10 class name if necessary
  if EditHandle = 0 then
    EditHandle := FindWindowEx(NotepadHandle, 0, 'Edit', nil);
  if EditHandle = 0 then
  begin
    ShowMessage('Windows 11 has a new rich edit Notepad that is not supported by ScriptPilot.');
    Exit;
  end;
  // Copy the text from the ValueListEditor to the clipboard
  Clipboard.AsText := ValueListEditor.Strings.Text;
  // Set focus to the Edit control in Notepad
  SetFocus(EditHandle);
  // Paste the text from the clipboard into the Edit control
  SendMessage(EditHandle, WM_PASTE, 0, 0);
  // Move the cursor to the beginning of the Edit control
  SendMessage(EditHandle, EM_SETSEL, 0, 0);
  // Scroll the document to the top
  SendMessage(EditHandle, EM_SCROLLCARET, 0, 0);
end;

procedure ModifyValueListEditorItem(editor: TValueListEditor; const keyName: string; const action: string; const newValue: string = '');
var
  rowIndex: Integer;
begin
  rowIndex := editor.Strings.IndexOfName(keyName);
  if rowIndex > -1 then
  begin
    if SameText(action, 'delete') then
    begin
      editor.Strings.Delete(rowIndex);
    end
    else if SameText(action, 'update') then
    begin
      editor.Strings.Values[keyName] := newValue;
    end;
  end;
end;

procedure LoadVariablesFromFile(const FileName: string);
var
  ExistingValues: TStringList;
  LoadedValues: TStringList;
  i: Integer;
begin
  ExistingValues := TStringList.Create;
  LoadedValues := TStringList.Create;
  try
    ExistingValues.Assign(frmMain.ValueListEditorVariables.Strings);
    LoadedValues.LoadFromFile(FileName);
    for i := 0 to LoadedValues.Count - 1 do
    begin
      if ExistingValues.IndexOfName(LoadedValues.Names[i]) = -1 then
      begin
        ExistingValues.Add(LoadedValues.Strings[i]);
      end;
    end;
    frmMain.ValueListEditorVariables.Strings.Assign(ExistingValues);
  finally
    ExistingValues.Free;
    LoadedValues.Free;
  end;
end;

Procedure SSHConnectEx(HostName : String; Username : String; Password : String);
var
  sshClient: TScSSHClient;
  sshShell: TScSSHShell;
  AreWeAtMax : Integer;
begin
AreWeAtMax := 1;
sshClient := TScSSHClient.Create(nil);
sshClient.KeyStorage := frmMain.ScMemoryStorage1;
sshClient.OnServerKeyValidate := frmMain.sshclient01ServerKeyValidate;
sshClient.OnBanner := frmMain.sshclient01Banner;
sshShell := TScSSHShell.Create(nil);
sshShell.Client := sshClient;
sshShell.OnAsyncReceive := frmMain.sshshell01AsyncReceive;
frmMain.SSHClientList.Add(sshClient);
frmMain.SSHShellList.Add(sshShell);
    while not sshClient.Connected do
      begin
        sshClient.HostName := HostName;
        sshClient.User := Username;
        sshClient.Password := Password;
          try
            sshClient.Connect;
          except
            on E: Exception do begin
                LogItStamp(E.Message,0);
                LogItStamp('No connection to SSH host ' + HostName + ' Retrying in ' + inttostr(Timeout) + ' seconds ' + E.Message,0);
                  if Pos('Authentication failed', E.Message) > 0 then begin
                  frmMain.SSHClientList.Remove(sshClient);
                  frmMain.SSHShellList.Remove(sshShell);
                  sshShell.Free;
                  sshClient.Free;
                  exit;
                  end;
              end;
          end;
      if sshClient.Connected then
        begin
          LogItStamp('SSH connected to host ' + HostName,0);
          sshShell.Connect;
          sshShell.WriteString('' + #13#10);
          Sleep(300);
        end;
    Inc(AreWeAtMax, Timeout);
      if AreWeAtMax > MaxTimeOut then
        begin
          LogItStamp('That was the last retry. Script will skip SSH connection to ' + HostName + ' and continue',0);
          sshShell.Free;
          sshClient.Free;
          frmMain.SSHClientList.Remove(sshClient);
          frmMain.SSHShellList.Remove(sshShell);
          exit;
        end;
    if not sshClient.Connected then Chill(Timeout);
  end;
end;

Procedure SSHCommandEx(HostName : String; Command : String);
var
  sshShell: TScSSHShell;
  sshClient: TScSSHClient;
  AreWeAtMax: Integer;
  i: Integer;
begin
  Verbose := true;
  AreWeAtMax := 0;
  for i := 0 to frmMain.SSHClientList.Count - 1 do
  begin
    sshClient := frmMain.SSHClientList[i];
    if sshClient.HostName = HostName then
    begin
      sshShell := frmMain.SSHShellList[i];
      while AreWeAtMax < MaxTimeOut do
      begin
        if not sshClient.Connected then
        begin
          LogItStamp('SSH not connected, trying to reconnect to host ' + HostName,0);
          try
            sshClient.Connect;
            if sshClient.Connected then
            begin
              LogItStamp('Reconnected to SSH host ' + HostName,0);
              sshShell.Connect;
            end
            else
            begin
              LogItStamp('Failed to reconnect. Waiting for ' + IntToStr(Timeout) + ' seconds before retrying connection to ' + HostName,0);
              Chill(Timeout);
              continue;
            end;
          except
            on E: Exception do
            begin
              LogItStamp('Error reconnecting to SSH host ' + HostName + ' - ' + E.Message + ' Retry: ' + IntToStr(AreWeAtMax + 1), 0);
                  if Pos('Authentication failed', E.Message) > 0 then begin
                  frmMain.SSHClientList.Remove(sshClient);
                  frmMain.SSHShellList.Remove(sshShell);
                  sshShell.Free;
                  sshClient.Free;
                  exit;
                  end;
            end;
          end;
        end;
        try
          sshShell.WriteString(Command + #13#10);
          LogItStamp('Sending SSH command "' + Command + '" to ' + HostName,0);
          Sleep(300);
          exit;
        except
          on E: Exception do
          begin
            LogItStamp('Error sending SSH command: ' + E.Message + ' Retry: ' + IntToStr(AreWeAtMax + 1), 0);
            sshClient.Connected := false;
          end;
        end;
        Inc(AreWeAtMax, Timeout);
      end;
      LogItStamp('Failed to send SSH command after ' + IntToStr(MaxTimeout) + ' retries', 0);
      exit;
    end;
  end;
 LogItStamp('Error: No active SSH connection for host ' + HostName + ' - Please use SSHConnect before calling SSHCommand for the first time to any host', 0);
 LogItStamp('Error: ScriptPilot can only send commands to hosts that first has been setup with SSHConnect', 0);
 LogItStamp('Error: When a host is in the active connection list, you can send as many commands as needed and ScriptPilot will also try to reconnect, if needed', 0);
end;

Procedure SSHDisconnectEx(HostName: string);
var
  i: Integer;
  sshClient: TScSSHClient;
  sshShell: TScSSHShell;
begin
  for i := 0 to frmMain.SSHClientList.Count - 1 do
  begin
    sshClient := frmMain.SSHClientList[i];
    sshShell := frmMain.SSHShellList[i];
    if sshClient.HostName = HostName then
    begin
      if sshClient.Connected then
      begin
        sshClient.Disconnect;
      end;
      sshShell.Free;
      frmMain.SSHShellList.Delete(i);
      sshClient.Free;
      frmMain.SSHClientList.Delete(i);
      LogItStamp('SSH disconnected and resources freed for host ' + HostName,0);
      exit;
    end;
  end;
  LogItStamp('No SSH client found for host ' + HostName, 0);
end;



end.
