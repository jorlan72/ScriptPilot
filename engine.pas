unit engine;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Menus, Vcl.ComCtrls,
  Vcl.StdCtrls, FlexCel.Core, FlexCel.XlsAdapter, ScSSHChannel, ScSSHClient, ScBridge,
  vcl.ValEdit, System.UITypes, system.DateUtils, IdException;

Procedure ProcessCommands;
Procedure RunLogit;
Procedure RunLogitStamp;
Procedure RunProgressBar;
procedure RunRunApp;
procedure RunMessageImage;
procedure RunAskForInputFields;
procedure RunInputBox;
procedure RunRunAppAndWait;
procedure RunConnectSSH;
procedure RunDisconnectSSH;
Procedure RunWaitForConnectionSSH;
Procedure RunWaitForSeconds;
Procedure RunSSHCommand;
Procedure RunSSHCommandVerbose;
Procedure RunWaitForFile;
Procedure RunGenerateNotepadOutput;
Procedure RunSSHContent;
Procedure RunFetchLogToVariable;
Procedure RunIfYes;
Procedure RunGotoRow;
Procedure RunExportLog;
Procedure RunquitScriptPilot;
Procedure RunHideScriptRow;
Procedure RunShowScriptRow;
Procedure RunVariableCreate;
Procedure RunVariableSet;
Procedure RunVariableIncrease;
Procedure RunVariableDecrease;
Procedure RunGotoIFVariableEquals;
Procedure RunWaitForPingReply;
Procedure Rungotoiffileexist;
Procedure RunGotoIFDay;
Procedure RunGotoIFWeek;
Procedure RunGotoIFMonth;
Procedure RunGotoIfHostAlive;
Procedure RunGetPublicIP;
Procedure RunDataSourceSetIndex;
Procedure RunGotoIfSSHConnected;
Procedure RunInfoMessage;
Procedure RunTelnetConnect;
Procedure RunTelnetDisConnect;
Procedure RunTelnetCommand;
Procedure RunTelnetCommandVerbose;
Procedure RunGotoIfTelnetConnected;
Procedure RunWaitForConnectionTelnet;
Procedure RunTelnetContent;
Procedure RunSSHTextFile;
Procedure RunTelnetTextFile;
Procedure RunGotoIFURL;
Procedure RunSendSyslog;
Procedure RunCMDCommand;
Procedure RunCMDCommandVerbose;
Procedure RunPSCommand;
Procedure RunPSCommandVerbose;
Procedure RunCMDContent;
Procedure RunCMDTextFile;
Procedure RunPSContent;
Procedure RunPSTextFile;
Procedure RunGotoSelector;
Procedure RunLoadVariables;
Procedure RunSaveVariables;
Procedure RunSSHConnectEx;
Procedure RunSSHCommandEx;
Procedure RunSSHDisconnectEx;

implementation

{$POINTERMATH ON}


uses
  main, snippets;


Procedure ProcessCommands;
var
 WasACommandHit : Boolean;
begin
WasACommandHit := false;

   if LowerCase(ExcelCommand) = 'logit' then begin
     RunLogit;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'logittime' then begin
     RunLogitStamp;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'progressbar' then begin
     RunProgressBar;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'remark' then begin
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'message' then begin
     if ExcelParam01 = '' then ExcelParam01 := 'No text in the message parameter cell';
     LogItStamp('Showing message to user and waiting for OK - "' + ExcelParam01 + '"',0);
     Showmessage('     ' + ExcelParam01 + '     ');
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'runapp' then begin
     RunRunApp;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'messageimage' then begin
     RunMessageImage;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'askforinputtable' then begin
     RunAskForInputFields;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'askforinput' then begin
     RunInputBox;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'runappandwait' then begin
     RunRunAppAndWait;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sshconnect' then begin
     RunConnectSSH;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sshdisconnect' then begin
     RunDisconnectSSH;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'waitforconnectionssh' then begin
     RunWaitForConnectionSSH;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'waitforseconds' then begin
     RunWaitForSeconds;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sshcommand' then begin
     RunSSHCommand;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sshcommandverbose' then begin
     RunSSHCommandVerbose;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'waitforfile' then begin
     RunWaitForFile;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'generatenotepadoutput' then begin
     RunGenerateNotepadOutput;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sshcontent' then begin
     RunSSHContent;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'fetchlogtovariable' then begin
     RunFetchLogToVariable;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoifyes' then begin
     RunIfYes;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotorow' then begin
     RunGotoRow;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'exportlog' then begin
     RunExportLog;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'quitscriptpilot' then begin
     Runquitscriptpilot;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'scriptrowoff' then begin
     RunHideScriptRow;
     WasACommandHit := true;
     if LowerCase(ExcelParam01) = 'silent' then exit;
     LogItStamp('Turning scriptrow off',0);
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'scriptrowon' then begin
     RunShowScriptRow;
     WasACommandHit := true;
     if LowerCase(ExcelParam01) = 'silent' then exit;
     LogItStamp('Turning scriptrow on',0);
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'variablecreate' then begin
     RunVariableCreate;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'variableset' then begin
     RunVariableSet;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'variableincrease' then begin
     RunVariableIncrease;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'variabledecrease' then begin
     RunVariableDecrease;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoifvariableequals' then begin
     RunGotoIFVariableEquals;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'telnetconnect' then begin
     RunTelnetConnect;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'waitforpingreply' then begin
     RunWaitForPingReply;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoiffileexist' then begin
     Rungotoiffileexist;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoifday' then begin
     RunGotoIfDay;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoifweek' then begin
     RunGotoIfWeek;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoifmonth' then begin
     RunGotoIfMonth;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoifhostalive' then begin
     RunGotoIfHostAlive;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'getpublicip' then begin
     RunGetPublicIP;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'trimoff' then begin
     WillTrim := false;
     WasACommandHit := true;
     if LowerCase(ExcelParam01) = 'silent' then exit;
     LogItStamp('Turning SSH and Telnet trimming off',0);
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'trimon' then begin
     WillTrim := true;
     WasACommandHit := true;
     if LowerCase(ExcelParam01) = 'silent' then exit;
     LogItStamp('Turning SSH and Telnet trimming on',0);
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'generatewordoutput' then begin
     LogItStamp('Starting Word document seacrh and replace',0);
     SearchAndReplaceWord(vartostr(ExcelParam01), vartostr(ExcelParam02));
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'datasourcesetindex' then begin
     RunDataSourceSetIndex;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoifsshconnected' then begin
     RunGotoIfSSHConnected;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = '#label#' then begin
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'clearlog' then begin
     if LowerCase(ExcelParam01) = 'export' then RunExportLog;
     frmMain.MemoLog.Clear;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'logiton' then begin
     WillLog := true;
     WasACommandHit := true;
     if LowerCase(ExcelParam01) = 'silent' then exit;
     LogItStamp('Turning logging on',0);
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'logitoff' then begin
     WillLog := false;
     WasACommandHit := true;
     if LowerCase(ExcelParam01) = 'silent' then exit;
     LogItStamp('Turning logging off',0);
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'informationon' then begin
     frmMain.groupboxinformation.visible := true;
     WasACommandHit := true;
     if LowerCase(ExcelParam01) = 'silent' then exit;
     LogItStamp('Turning information banner on',0);
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'informationoff' then begin
     frmMain.groupboxinformation.visible := false;
     WasACommandHit := true;
     if LowerCase(ExcelParam01) = 'silent' then exit;
     LogItStamp('Turning information banner off',0);
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'infomessage' then begin
     RunInfoMessage;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'telnetdisconnect' then begin
     RunTelnetDisConnect;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'telnetcommand' then begin
     RunTelnetCommand;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'telnetcommandverbose' then begin
     RunTelnetCommandVerbose;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoiftelnetconnected' then begin
     RunGotoIfTelnetConnected;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'waitforconnectiontelnet' then begin
     RunWaitForConnectionTelnet;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'telnetcontent' then begin
     RunTelnetContent;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'telnettextfile' then begin
     RunTelnetTextFile;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sshtextfile' then begin
     RunSSHTextFile;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoifurl' then begin
     RunGotoIFURL;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sendsyslog' then begin
     RunSendSysLog;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'cmdcommand' then begin
     RunCMDCommand;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'cmdcommandverbose' then begin
     RunCMDCommandVerbose;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'pscommand' then begin
     RunPSCommand;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'pscommandverbose' then begin
     RunPSCommandVerbose;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'cmdcontent' then begin
     RunCMDContent;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'cmdtextfile' then begin
     RunCMDTextFile;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'pscontent' then begin
     RunPSContent;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'pstextfile' then begin
     RunPSTextFile;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'gotoselector' then begin
     RunGotoSelector;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'loadvariables' then begin
     RunLoadVariables;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'savevariables' then begin
     RunSaveVariables;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sshconnectex' then begin
     RunSSHConnectEx;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sshcommandex' then begin
     RunSSHCommandEx;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

   if LowerCase(ExcelCommand) = 'sshdisconnectex' then begin
     RunSSHDisconnectEx;
     WasACommandHit := true;
   end;
   if WasACommandHit then exit;

if not WasACommandHit then begin
  LogItStamp('ERROR: An unknown command was found in the script at row ' + inttostr(CurrentRow),0);
  LogItStamp('Command not recognized was : ' + vartostr(ExcelCommand) + ' : ' + vartostr(ExcelParam01) + ' : ' + vartostr(ExcelParam02) + ' : ' + vartostr(ExcelParam03),0);
end;

end;


Procedure RunLogit;
var
 linestoadd : Integer;
begin
if ExcelParam02 = '' then
   begin
    linestoadd := 0;
    LogIt(ExcelParam01, linestoadd);
    exit;
   end;
 if not trystrtoint(ExcelParam02, linestoadd) then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The cell is either empty or not a number - Setting value to default "0"',0);
    linestoadd := 0;
    end else begin
              if linestoadd < 0 then
                begin
                LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
                LogItStamp('The cell is a number lower than 0 - Setting value to default "0"',0);
                linestoadd := 0;
                end;
             end;
LogIt(ExcelParam01, linestoadd);
end;

Procedure RunLogitStamp;
var
 linestoadd : Integer;
begin
if ExcelParam02 = '' then
   begin
    linestoadd := 0;
    LogItStamp(ExcelParam01, linestoadd);
    Exit;
   end;
 if not trystrtoint(ExcelParam02, linestoadd) then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The cell is either empty or not a number - Setting value to default "0"',0);
    linestoadd := 0;
    end else begin
              if linestoadd < 0 then
                begin
                LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
                LogItStamp('The cell is a number lower than 0 - Setting value to default "0"',0);
                linestoadd := 0;
                end;
             end;
LogItStamp(ExcelParam01, linestoadd);
end;

Procedure RunProgressBar;
var
 bar : integer;
begin
  if not trystrtoint(ExcelParam01, bar) then begin
  LogItStamp('ERROR: An error was found for the progressbar value setting at row ' + inttostr(CurrentRow),0);
  LogItStamp('The cell is either empty or not a number',0);
  end else begin
             if bar > 100 then begin
                                LogItStamp('ERROR: An error was found for the progressbar value setting at row ' + inttostr(CurrentRow),0);
                                LogItStamp('The cell is a number higher than 100',0);
                                exit;
                               end;
             if bar < 0 then begin
                              LogItStamp('ERROR: An error was found for the progressbar value setting at row ' + inttostr(CurrentRow),0);
                              LogItStamp('The cell is a number lower than 0',0);
                              exit;
                             end;
           frmMain.ProgressBar.Position := bar;
           LogItStamp('Progressbar setting ' + inttostr(bar) + '%',0);
           end;

end;

procedure RunRunApp;
begin
  if FindExecutable(ExcelParam01) <> '' then begin
    if ExcelParam02 <> '' then LogItStamp('Starting external application ' + ExcelParam01 + ' with parameter ' + ExcelParam02,0);
    if ExcelParam02 = '' then LogItStamp('Starting external application ' + ExcelParam01,0);
    ShellOpen(ExcelParam01,ExcelParam02);
  end else begin
    LogItStamp('ERROR: The external application ' + ExcelParam01 + ' could not be found.',0);
    LogItStamp('Skipping command at row ' + inttostr(CurrentRow),0);
  end;
end;

procedure RunMessageImage;
begin
 frmMain.LabelMessageImage.Caption := vartostr(ExcelParam01);
 frmMain.ButtonMessageImageContinue.Caption := vartostr(ExcelParam03);
 frmMain.CardPanelRight.ActiveCard := frmmain.CardMessageImage;
 if fileexists(extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\data\' + vartostr(excelparam02)) then begin
   LogItStamp('Showing message "' + vartostr(ExcelParam01) + '" to user with picture "' + vartostr(excelparam02 + '"'),0);
   frmmain.MessageImage.Picture.LoadFromFile(extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\data\' + vartostr(excelparam02));
   LogItStamp('Waiting for user to continue',0);
 end else begin
   LogItStamp('WARNING: Showing message "' + vartostr(ExcelParam01) + '" to user, but picture "' + vartostr(excelparam02) + '" was not found',0);
   LogItStamp('Waiting for user to continue',0);
 end;
WaitForContinue;
if not stoppressed then LogItStamp('User clicked to continue script',0);
frmMain.CardPanelRight.ActiveCard := frmMain.CardLog;
end;

procedure RunAskForInputFields;
var
 Counter: Integer;
 keyindex : integer;
 keyname : string;
begin
 frmMain.LabelAskForInput.caption := vartostr(ExcelParam01);
 frmMain.CardPanelRight.ActiveCard := frmMain.CardVariableInput;
 LogItStamp('Asking user for variable input taken from the settings sheet',0);
 LogItStamp('Waiting for user to continue',0);
 WaitForContinue;
   if not stoppressed then LogItStamp('User clicked to continue script',0);
      if not stoppressed then begin
      LogItStamp('Adding user input variables to the global variables for the duration of the script',0);
         for Counter := 0 to frmMain.ValueListEditorInput.strings.Count - 1 do begin
           keyname := InputArray[Counter].VariableName;
           keyindex := FindValueListEditorKey(frmMain.ValueListEditorVariables, keyname);
             if keyindex = -1 then frmMain.ValueListEditorVariables.InsertRow(InputArray[Counter].VariableName, frmMain.ValueListEditorInput.Strings.ValueFromIndex[Counter], true)
                              else LogItStamp('Skipping add variable ' + keyname + ' as it is already in the global variable table',0);
         end;
      end;
 frmMain.CardPanelRight.ActiveCard := frmMain.CardLog;
end;

procedure RunInputBox;
var
 inputvariable : string;
 inputtitle : string;
 defaultvalue : string;
 inputreply : string;
 keyindex : integer;
 keyname : string;
begin
 inputvariable := vartostr(ExcelParam01);
 keyname := inputvariable;
 keyindex := FindValueListEditorKey(frmMain.ValueListEditorVariables, keyname);
 if keyindex <> -1 then
    begin
      LogItStamp('ERROR: Will not create the variable ' + keyname + ', as it already exists. Skipping command at row ' + inttostr(CurrentRow),0);
      Exit
    end;
 inputtitle := vartostr(ExcelParam02);
 defaultvalue := vartostr(ExcelParam03);
 LogItStamp('Asking user for input "' + inputtitle + '" and assigning input to variable "' + inputvariable + '"',0);
 inputreply := defaultvalue;
 inputreply := InputBox(inputtitle, 'Please enter your input :', defaultvalue);
 LogItStamp('Variable "' + inputvariable + '" is assigned the value "' + inputreply + '"',0);
 LogItStamp('Adding user input variable to the global variables for the duration of the script',0);
 frmMain.ValueListEditorVariables.InsertRow(inputvariable, inputreply, true);
end;

procedure RunRunAppAndWait;
begin
 if FindExecutable(ExcelParam01) <> '' then begin
    if ExcelParam02 <> '' then LogItStamp('Starting external application ' + ExcelParam01 + ' with parameter ' + ExcelParam02 + ' and pausing script until application is closed',0);
    if ExcelParam02 = '' then LogItStamp('Starting external application ' + ExcelParam01 + ' and pausing script until application is closed',0);
    ExternalProcess := ShellExecuteAndWait(ExcelParam01,ExcelParam02);
 end else begin
    LogItStamp('ERROR: The external application ' + ExcelParam01 + ' could not be found',0);
    LogItStamp('Skipping command at row ' + inttostr(CurrentRow),0);
 end;
end;

procedure RunConnectSSH;
 begin
   if frmMain.sshclient01.Connected then LogItStamp('SSH connection to another host already active. Ignoring command at row ' + inttostr(CurrentRow),0);
   if frmMain.sshclient01.Connected then exit;
     LogItStamp('Trying SSH connection to host ' + ExcelParam01 + ' with user ' + ExcelParam02 + ' and timeout of ' + inttostr(ChillTime) + ' seconds',0);
     frmMain.TimerSSH.Enabled := true;
     frmMain.sshclient01.HostName := VarToStr(ExcelParam01);
     frmMain.sshclient01.User := VarToStr(ExcelParam02);
     frmMain.sshclient01.Password := VarToStr(ExcelParam03);
       try
         try
         frmMain.sshclient01.Connect;
         except
         on E: Exception do
         begin
           LogItStamp('ERROR: Could not connect to SSH host ' + VarToStr(ExcelParam01) + ' - If welcome banner received, check username and password', 0);
         end;
         end;
       finally
         frmMain.TimerSSH.Enabled := false;
       end;
     if frmMain.sshclient01.Connected then frmMain.sshshell01.Connect;
     if frmMain.sshclient01.connected then frmMain.sshshell01.WriteString('' + #13#10);
     if frmMain.sshclient01.Connected then LogItStamp('SSH connected to host ' + ExcelParam01 + ' with user ' + ExcelParam02,0);
     if frmMain.sshclient01.Connected then Chill(ChillTime);
     sshinputstring := '';
end;

procedure RunDisconnectSSH;
begin
 frmmain.sshclient01.Disconnect;
 frmMain.sshshell01.Disconnect;
 LogItStamp('Disconnecting SSH',0);
end;

Procedure RunWaitForConnectionSSH;
var
 AreWeAtMax : Integer;
begin
 AreWeAtMax := 0;
   if frmMain.sshclient01.Connected then LogItStamp('ERROR: SSH connection to another host already active. Ignoring command at row ' + inttostr(CurrentRow),0);
   if frmMain.sshclient01.Connected then exit;
   LogItStamp('Trying SSH connection to host ' + ExcelParam01 + ' with loop interval of ' + inttostr(timeout) + ' seconds',0);
   LogItStamp('Click the Stop button to abort',0);
   if MaxTimeout > 0 then LogItStamp('Max retry time set to ' + inttostr(MaxTimeout) + ' seconds',0);
         if MaxTimeout <= Timeout then begin
           LogItStamp('ERROR: Max retry time set lower than the connection time in settings. Adjusting to 180 seconds',0);
           MaxTimeout := 180;
         end;
    while not frmMain.sshclient01.Connected do begin
       if StopPressed then exit;
         frmMain.TimerSSH.Enabled := true;
         frmMain.sshclient01.HostName := VarToStr(ExcelParam01);
         frmMain.sshclient01.User := VarToStr(ExcelParam02);
         frmMain.sshclient01.Password := VarToStr(ExcelParam03);
        try
           try
            frmMain.sshclient01.Connect;
           except
            on E: Exception do
            begin
              LogItStamp('SSH host not ready yet. Retrying in ' + inttostr(timeout) + ' seconds - If welcome banner received, check username and password',0);
            end;
           end;
         finally
         frmMain.TimerSSH.Enabled := false;
        end;
       if frmMain.sshclient01.Connected then LogItStamp('SSH connected to host ' + ExcelParam01 + ' with user ' + ExcelParam02,0);
       if frmMain.sshclient01.Connected then frmMain.sshshell01.Connect;
       if frmMain.sshclient01.connected then frmMain.sshshell01.WriteString('' + #13#10);
       if frmMain.sshclient01.Connected then Chill(ChillTime);
       Inc(AreWeAtMax, Timeout);
         if AreWeAtMax >= MaxTimeout then begin
          LogItStamp('Will not try again. Max retry time reached. Script will continue',0);
          exit;
         end;
       if not frmMain.sshclient01.Connected then Chill(Timeout);
    end;
end;

Procedure RunWaitForSeconds;
var
 secondstowait : integer;
begin
 if not trystrtoint(ExcelParam01, secondstowait) then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The cell is either empty or not a number',0);
    end else begin
              if secondstowait < 1 then
                begin
                LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
                LogItStamp('The cell is a number lower than 1',0);
                exit;
                end;
              LogItStamp('Pausing script for ' + inttostr(secondstowait) + ' seconds',0);
              Chill(secondstowait);
              LogItStamp('Starting script again after a ' + inttostr(secondstowait) + ' second wait',0);
             end;
end;

Procedure RunSSHCommand;
begin
verbose := false;
  if not frmMain.sshclient01.Connected then begin
   LogItStamp('ERROR: Skipping command at row ' + inttostr(CurrentRow) + '. SSH not connected',0);
  end else begin
   frmMain.sshshell01.WriteString(vartostr(ExcelParam01) + #13#10);
   LogItStamp('Sending SSH command "' + vartostr(ExcelParam01) + '" to connected host',0);
   chill(ChillTime);
  end;
end;

Procedure RunSSHCommandVerbose;
begin
  if not frmMain.sshclient01.Connected then begin
   LogItStamp('ERROR: Skipping command at row ' + inttostr(CurrentRow) + '. SSH not connected',0);
  end else begin
   Verbose := true;
   LastVerboseStart := frmMain.MemoLog.Lines.Count - 1;
   LogItStamp('Sending SSH Verbose command "' + vartostr(ExcelParam01) + '" to connected host',0);
   LogItStamp('Displaying data returned from host:',1);
   sshinputstring := '';
   frmMain.sshshell01.WriteString(vartostr(ExcelParam01) + #13#10);
   chill(ChillTime);
  end;
Verbose := false;
LogIt('',0);
end;

Procedure RunWaitForFile;
var
 WaitTimeOut : Integer;
 WasFileFound : Boolean;
 stopthisshit : Boolean;
begin
stopthisshit := false;
WasFileFound := false;
 if not trystrtoint(ExcelParam03, WaitTimeOut) then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The cell is either empty or not a number',0);
    StopThisShit := true;
   end else begin
               if WaitTimeout < 0 then
                begin
                LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
                LogItStamp('The cell is a number lower than 0',0);
                StopThisShit := true;
                end;
            end;
if stopthisshit then exit;
if ExcelParam02 = '' then ExcelParam02 := 'No-File-Specified.txt';
if ExcelParam02 = 'No-File-Specified.txt' then WaitTimeOut := 10;
    if not DirectoryExists(vartostr(ExcelParam01)) then
    begin
     LogItStamp('ERROR: Can not find the search directory',0);
    end else
      begin
      LogItStamp('Waiting for file ' + vartostr(ExcelParam02) + ' in directory ' + vartostr(ExcelParam01),0);
        if WaitTimeOut > 0 then LogItStamp('Will wait for ' + inttostr(WaitTimeOut) + ' seconds',0) else
         LogItStamp('Will wait for forever until file is found or you hit the Stop button',0);
         WasFileFound := WaitForFile(vartostr(ExcelParam01), vartostr(ExcelParam02), WaitTimeOut);
      end;
 if WasFileFound then LogItStamp('File was found.',0) else LogItStamp('File NOT found.',0);
end;

Procedure RunGenerateNotepadOutput;
var
 xls : TXlsFile;
 Counter : integer;
 FileName : string;
begin
if ExcelParam01 = '' then begin
  LogItStamp('ERROR: Generating Notepad document aborted. No content sheet specified',0);
  exit
end;
  LogItStamp('Generating Notepad document from content sheet ' + vartostr(ExcelParam01),0);
  xls := TXlsFile.Create;
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.xlsx';
  CopyFileToTemp(filename);
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.tmp';
      if SheetExists(filename,vartostr(ExcelParam01)) then begin
          xls.open(filename);
          xls.ActiveSheetByName := vartostr(ExcelParam01);
          Counter := 1;
          frmMain.MemoContent.clear;
                while not xls.GetCellValue(counter,1).IsEmpty do begin
                  frmMain.MemoContent.lines.add(vartostr(xls.GetCellValue(counter,1)));
                  inc(counter);
                  if StopPressed then exit;
                  application.ProcessMessages;
                end;
               xls.free;
          for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do begin
            frmMain.MemoContent.text := StringReplace(frmMain.MemoContent.text, frmMain.ValueListEditorVariables.Strings.KeyNames[Counter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter], [rfReplaceAll]);
            if StopPressed then exit;
            application.ProcessMessages;
          end;
          OpenNotepadWithMemoText(frmMain.MemoContent);
          DeleteFile(filename);
      end else begin
        LogItStamp('ERROR: Content sheet not found. No content to work with. Aborting.',0);
      end;
end;

Procedure RunSSHContent;
var
 xls : TXlsFile;
 Counter : integer;
 FileName : string;
 WaitTimeOut : integer;
 PushMemoLine : Integer;
 ChillCounter : Integer;
 NoFeedbackCounter : Integer;
 DelayedCounter : Integer;
 Delayed : Boolean;
begin
if not trystrtoint(ExcelParam02, WaitTimeOut) then WaitTimeOut := 300;
if ExcelParam01 = '' then begin
  LogItStamp('ERROR: Sending SSH content aborted. No content sheet specified',0);
  exit
end;
  LogItStamp('Getting SSH content ready from content sheet ' + vartostr(ExcelParam01),0);
  xls := TXlsFile.Create;
  DelayedCounter := 0;
  NoFeedbackCounter := 0;
  Delayed := false;
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.xlsx';
  CopyFileToTemp(filename);
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.tmp';
      if SheetExists(filename,vartostr(ExcelParam01)) then begin
          xls.open(filename);
          xls.ActiveSheetByName := vartostr(ExcelParam01);
          Counter := 1;
          frmMain.MemoContent.clear;
                while not xls.GetCellValue(counter,1).IsEmpty do begin
                  frmMain.MemoContent.lines.add(vartostr(xls.GetCellValue(counter,1)));
                  inc(counter);
                  if StopPressed then exit;
                  application.ProcessMessages;
                end;
               xls.free;
          for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do begin
            frmMain.MemoContent.text := StringReplace(frmMain.MemoContent.text, frmMain.ValueListEditorVariables.Strings.KeyNames[Counter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter], [rfReplaceAll]);
            if StopPressed then exit;
            application.ProcessMessages;
          end;
             if not frmMain.sshclient01.Connected then LogItStamp('ERROR: Skipping SSH Content push. SSH not connected',0);
              if frmMain.sshclient01.Connected then begin
               PushMemoLine := frmMain.MemoLog.lines.Count;
               LogItStamp('Update log line # ' + inttostr(PushMemoLine),0);
               LogItStamp('Start pushing content to host from sheet ' + vartostr(ExcelParam01) + ' containing ' + inttostr(frmMain.MemoContent.Lines.Count) + ' lines. Line send delay is ' + inttostr(WaitTimeOut) + ' milliseconds',1);
               PushMemoLine := frmMain.MemoLog.lines.Count;
                 for counter := 0 to frmMain.MemoContent.Lines.Count -1 do begin
                 feedbackreceived := false;
                 frmMain.sshshell01.WriteString(frmMain.MemoContent.lines.Strings[Counter] + #13#10);
                 sleep(WaitTimeOut);
                 application.ProcessMessages;
                   for chillcounter := 1 to 10 do begin
                     if not feedbackreceived then chill(1);
                     if not feedbackreceived then delayed := true;
                   end;
                 if delayed then inc(DelayedCounter);
                 if not feedbackreceived then inc(NoFeedbackCounter);
                 frmMain.MemoLog.lines[PushMemoLine] := '                      Written ' + inttostr(counter + 1) + ' of ' + inttostr(frmMain.MemoContent.Lines.Count) + ' lines to SHH host - Delayed transactions : ' + inttostr(DelayedCounter) + ' - No feedback from host : ' + inttostr(NoFeedbackCounter);
                 delayed := false;
                 if StopPressed then exit;
                 end;
                LogItStamp('SSH content push done.',0);
                DeleteFile(filename);
              end;
      end else begin
        LogItStamp('ERROR: Content sheet not found. No content to work with. Aborting.',0);
      end;
end;

Procedure RunFetchLogToVariable;
var
 HowMany : Integer;
 WordFound : String;
 keyindex : integer;
 keyname : string;
begin
 if not trystrtoint(ExcelParam03, HowMany) then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The cell is either empty or not a number',0);
   end else
     begin
      if Howmany < 1 then
        begin
         LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
         LogItStamp('The cell is a number lower than 1',0);
         exit;
        end;
           WordFound := FetchTextAfterSearch(frmMain.MemoLog, LastVerboseStart, frmMain.MemoLog.Lines.Count -1, vartostr(ExcelParam02), HowMany);
               if WordFound <> '' then
                begin
                 LogItStamp('Grabbed the text ' + WordFound + ' following the search string ' + vartostr(ExcelParam02),0);
                end else
                   begin
                    LogItStamp('ERROR: Did not find the search string ' + vartostr(ExcelParam02) + ' - Skipping text grab from the log',0)
                   end;
     end;
 if WordFound <> '' then
 begin
  keyname := vartostr(ExcelParam01);
  keyindex := FindValueListEditorKey(frmMain.ValueListEditorVariables, keyname);
  if keyindex = -1 then frmMain.ValueListEditorVariables.Strings.AddPair(vartostr(ExcelParam01), WordFound);
  if keyindex = -1 then  LogItStamp('Grabbed text put it into variable ' + vartostr(ExcelParam01),0);
   if keyindex <> -1 then
     begin
     frmMain.ValueListEditorVariables.Values[vartostr(ExcelParam01)] := WordFound;
     LogItStamp('Variable ' + vartostr(ExcelParam01) + ' updated with new value ' + WordFound,0);
     end;
 end;
end;

Procedure RunIfYes;
var
 UserChoice: Integer;
 RowNumber : Integer;
begin
 if not trystrtoint(ExcelParam02, RowNumber) then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The cell is either empty or not a number',0);
    end else begin
               if Rownumber < 2 then
                begin
                LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
                LogItStamp('The cell is a number lower than 2 and that reference is before script start',0);
                exit;
                end;
             end;
 LogItStamp('Asking the user : "' + vartostr(ExcelParam01) + '" with a "YES" or "NO" option',0);
 UserChoice := MessageDlg('          ' + vartostr(ExcelParam01) + '          ', mtConfirmation, [mbYes, mbNo], 0);
   if UserChoice = mrYes then
     begin
      LogItStamp('User answered "YES", and the script will jump to row ' + inttostr(RowNumber),0);
      CurrentRow := RowNumber - 1;
     end else
     begin
      LogItStamp('User answered "NO", and the script will continue at next row - Row : ' + inttostr(currentrow + 1),0);
     end;
end;

Procedure RunGotoRow;
var
 RowNumber : Integer;
begin
 if not trystrtoint(ExcelParam01, RowNumber) then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The cell is either empty or not a number',0);
    end else begin
              if Rownumber < 2 then
                begin
                LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
                LogItStamp('The cell is a number lower than 2 and that reference is before script start',0);
                exit;
                end;
              CurrentRow := RowNumber - 1;
              LogItStamp('Script jump to row ' + inttostr(RowNumber),0);
             end;
end;

Procedure RunExportLog;
var
 FileName : string;
 DateString : String;
begin
 FileName := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\ScriptPilot Log Export ';
 DateString := datetimetostr(now);
 DateString := StringReplace(DateString, ':', '.', [rfReplaceAll]);
 Filename := Filename + DateString + '.txt';
 frmMain.MemoLog.Lines.SaveToFile(Filename);
 LogItStamp('Exporting current log to : ' + FileName,0);
end;

Procedure Runquitscriptpilot;
begin
 QuitApplication := true;
end;

Procedure RunHideScriptRow;
begin
  frmMain.LabelScriptActiveRow.Visible := false;
end;

Procedure RunShowScriptRow;
begin
  frmMain.LabelScriptActiveRow.Visible := true;
end;

Procedure RunVariableCreate;
var
 keyindex : integer;
 keyname : string;
begin
  keyname := vartostr(ExcelParam01);
  keyindex := FindValueListEditorKey(frmMain.ValueListEditorVariables, keyname);
     if keyindex = -1 then
     begin
      LogItStamp('Creatng new variable named ' + vartostr(ExcelParam01) + ' with initial value "' + vartostr(ExcelParam02) + '"',0);
      frmMain.ValueListEditorVariables.Strings.AddPair(vartostr(ExcelParam01), vartostr(ExcelParam02));
     end else
       begin
        LogItStamp('ERROR: Will not create the variable ' + keyname + ', as it already exists. Skipping command at row ' + inttostr(CurrentRow),0);
       end;
end;

Procedure RunVariableSet;
var
 keyindex : integer;
 keyname : string;
begin
  keyname := vartostr(ExcelParam01);
  keyindex := FindValueListEditorKey(frmMain.ValueListEditorVariables, keyname);
    if keyindex = -1 then
    begin
     LogItStamp('ERROR: The variable ' + keyname + ' does not exist. Cannot set new value. Skipping command at row ' + inttostr(CurrentRow),0);
    end else
        begin
         frmMain.ValueListEditorVariables.Values[vartostr(ExcelParam01)] := vartostr(ExcelParam02);
         LogItStamp('Variable ' + keyname + ' updated with new value "' + vartostr(ExcelParam02) + '"',0);
        end;
end;

Procedure RunVariableIncrease;
var
 keyindex : integer;
 keyname : string;
 oldvalue, newvalue, sum : integer;
 check : integer;
begin
  sum := 0;
  check := 0;
  keyname := vartostr(ExcelParam01);
  keyindex := FindValueListEditorKey(frmMain.ValueListEditorVariables, keyname);
    if keyindex = -1 then
    begin
     LogItStamp('ERROR: The variable ' + keyname + ' does not exist. Cannot increase value. Skipping command at row ' + inttostr(CurrentRow),0);
    end else
        begin
           if trystrtoint(ExcelParam02, newvalue) then inc(check) else LogItStamp('ERROR: Script parameter does not contain a number in row ' + inttostr(CurrentRow),0);
           if trystrtoint(frmMain.ValueListEditorVariables.strings.ValueFromIndex[keyindex], oldvalue) then inc(check) else LogItStamp('ERROR: The variable to increase is not a number. Skipping command at row ' + inttostr(CurrentRow),0);
           sum := oldvalue + newvalue;
        end;
  if check = 2 then
  begin
   frmMain.ValueListEditorVariables.Values[vartostr(ExcelParam01)] := inttostr(sum);
   LogItStamp('Variable ' + keyname + ' increased to new value ' + inttostr(sum),0);
   end;
end;

Procedure RunVariableDecrease;
var
 keyindex : integer;
 keyname : string;
 oldvalue, newvalue, sum : integer;
 check : integer;
begin
  sum := 0;
  check := 0;
  keyname := vartostr(ExcelParam01);
  keyindex := FindValueListEditorKey(frmMain.ValueListEditorVariables, keyname);
    if keyindex = -1 then
    begin
     LogItStamp('ERROR: The variable ' + keyname + ' does not exist. Cannot decrease value. Skipping command at row ' + inttostr(CurrentRow),0);
    end else
        begin
           if trystrtoint(ExcelParam02, newvalue) then inc(check) else LogItStamp('ERROR: Script parameter does not contain a number in row ' + inttostr(CurrentRow),0);
           if trystrtoint(frmMain.ValueListEditorVariables.strings.ValueFromIndex[keyindex], oldvalue) then inc(check) else LogItStamp('ERROR: The variable to decrease is not a number. Skipping command at row ' + inttostr(CurrentRow),0);
           sum := oldvalue - newvalue;
        end;
 if check < 2 then exit;
     if sum < 2 then
       begin
        LogItStamp('ERROR: Decreasing the variables lower that 2 (script start). Skipping command at row ' + inttostr(CurrentRow),0);
        check := 0;
       end;
  if check = 2 then
  begin
   frmMain.ValueListEditorVariables.Values[vartostr(ExcelParam01)] := inttostr(oldvalue + newvalue);
   LogItStamp('Variable ' + keyname + ' decreased to new value ' + inttostr(sum),0);
   end;
end;

Procedure RunGotoIFVariableEquals;
var
 keyindex : integer;
 keyname : string;
 gotorow : integer;
begin
 keyname := vartostr(ExcelParam01);
  keyindex := FindValueListEditorKey(frmMain.ValueListEditorVariables, keyname);
    if keyindex = -1 then
    begin
     LogItStamp('ERROR: The variable ' + keyname + ' does not exist. Cannot compare values for goto. Skipping command at row ' + inttostr(CurrentRow),0);
     exit;
    end else
        begin
           if not trystrtoint(ExcelParam03, gotorow) then
             begin
             LogItStamp('ERROR: Script parameter does not contain a number for goto value in row ' + inttostr(CurrentRow),0);
             exit
             end;
        end;
if gotorow < 2 then
 begin
  LogItStamp('ERROR: Cannot goto any line before script start. Skipping command at row ' + inttostr(CurrentRow),0);
  exit;
 end;
if frmMain.ValueListEditorVariables.strings.ValueFromIndex[keyindex] = vartostr(ExcelParam02) then
   begin
     LogItStamp('Variables match - Jumping to line ' + vartostr(ExcelParam03),0);
     CurrentRow := gotorow - 1;
   end else
   begin
     LogItStamp('Variables does not match - staying put',0);
   end;
end;

Procedure RunWaitForPingReply;
var
 AreWeAtMax : Integer;
 DoThisPing : String;
begin
 AreWeAtMax := 0;
 GotReplyFromPing := false;
 PingHost := VarToStr(ExcelParam01);
 if not IsValidIPv4Address(vartostr(ExcelParam01)) then
         begin
          if GetHostIP(vartostr(ExcelParam01)) = '' then
            begin
              LogItStamp('ERROR: Not a valid IP address and could not resolve hostname. Row ' + inttostr(CurrentRow),0);
              exit
            end;
         end;
 DoThisPing := GetHostIP(PingHost);
 PingHost := DoThisPing;
   LogItStamp('Trying PING to host (' + ExcelParam01 + ') - ' + Pinghost + ' with loop interval of ' + inttostr(timeout) + ' seconds',0);
   LogItStamp('Click the Stop button to abort',0);
   if MaxTimeout > 0 then LogItStamp('Max retry time set to ' + inttostr(MaxTimeout) + ' seconds',0);
      if MaxTimeout <= Timeout then begin
       LogItStamp('ERROR: Max retry time set lower than the connection time in settings. Adjusting to 180 seconds',0);
       MaxTimeout := 180;
      end;
    while not GotReplyFromPing do begin
       if StopPressed then exit;
         frmMain.Ping.PingHost(PingHost);
         Chill(ChillTime);
           if not GotReplyFromPing then LogItStamp('Host not ready yet. Retrying in ' + inttostr(timeout) + ' seconds',0);
            Inc(AreWeAtMax, Timeout);
            if AreWeAtMax >= MaxTimeout then begin
              LogItStamp('Will not try again. Max retry time reached. Script will continue',0);
              exit;
            end;
       if not GotReplyFromPing then Chill(TimeOut);
    end;
  if GotReplyFromPing then LogItStamp('Received reply from host. Script will continue',0);
end;

Procedure Rungotoiffileexist;
var
 gotorow : integer;
begin
 if not trystrtoint(ExcelParam02, gotorow) then
        begin
         LogItStamp('ERROR: Script parameter does not contain a number for goto value in row ' + inttostr(CurrentRow),0);
         exit
        end;
 if gotorow < 2 then
        begin
         LogItStamp('ERROR: Cannot goto any line before script start. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if fileexists(vartostr(ExcelParam01)) then
        begin
         LogItStamp('File exist - Jumping to line ' + inttostr(gotorow),0);
         CurrentRow := gotorow - 1;
        end else begin
                  LogItStamp('File does not exist - Staying put',0);
                 end;
end;

Procedure RunGotoIFDay;
var
 gotorow : integer;
 CurrentDayInt: Integer;
 Selected : Integer;
begin
 CurrentDayInt := DayOfWeek(Now);
 if not trystrtoint(ExcelParam02, gotorow) then
        begin
         LogItStamp('ERROR: Script parameter does not contain a number for goto value in row ' + inttostr(CurrentRow),0);
         exit
        end;
 if gotorow < 2 then
        begin
         LogItStamp('ERROR: Cannot goto any line before script start. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if not trystrtoint(ExcelParam01, selected) then
        begin
         LogItStamp('ERROR: The day to match is not a number. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if selected > 7 then
        begin
         LogItStamp('ERROR: The day to match is a number higher than 7. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if selected < 1 then
        begin
         LogItStamp('ERROR: The day to match is a number lower than 1. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if selected = CurrentDayInt then
        begin
         LogItStamp('The weekday is a match! - Jumping to line ' + inttostr(gotorow),0);
         CurrentRow := gotorow - 1;
        end else
        begin
         LogItStamp('The weekday is NOT a match - Staying put',0);
        end;
end;

Procedure RunGotoIFWeek;
var
 gotorow : integer;
 CurrentWeek : Word;
 CurrentDate: TDateTime;
 Selected : Integer;
begin
 CurrentDate := Now;
 CurrentWeek := WeekOfTheYear(CurrentDate);
 if not trystrtoint(ExcelParam02, gotorow) then
        begin
         LogItStamp('ERROR: Script parameter does not contain a number for goto value in row ' + inttostr(CurrentRow),0);
         exit
        end;
 if gotorow < 2 then
        begin
         LogItStamp('ERROR: Cannot goto any line before script start. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if not trystrtoint(ExcelParam01, selected) then
        begin
         LogItStamp('ERROR: The day to match is not a number. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if selected > 52 then
        begin
         LogItStamp('ERROR: The day to match is a number higher than 52. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if selected < 1 then
        begin
         LogItStamp('ERROR: The day to match is a number lower than 1. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if selected = CurrentWeek then
        begin
         LogItStamp('The week is a match! - Jumping to line ' + inttostr(gotorow),0);
         CurrentRow := gotorow - 1;
        end else
        begin
         LogItStamp('The week is NOT a match - Staying put',0);
        end;
end;

Procedure RunGotoIFMonth;
var
 gotorow : integer;
 CurrentMonth: Word;
 CurrentDate: TDateTime;
 Selected : Integer;
begin
 CurrentDate := Now;
 CurrentMonth := MonthOf(CurrentDate);
 if not trystrtoint(ExcelParam02, gotorow) then
        begin
         LogItStamp('ERROR: Script parameter does not contain a number for goto value in row ' + inttostr(CurrentRow),0);
         exit
        end;
 if gotorow < 2 then
        begin
         LogItStamp('ERROR: Cannot goto any line before script start. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if not trystrtoint(ExcelParam01, selected) then
        begin
         LogItStamp('ERROR: The day to match is not a number. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if selected > 12 then
        begin
         LogItStamp('ERROR: The day to match is a number higher than 12. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if selected < 1 then
        begin
         LogItStamp('ERROR: The day to match is a number lower than 1. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if selected = CurrentMonth then
        begin
         LogItStamp('The month is a match! - Jumping to line ' + inttostr(gotorow),0);
         CurrentRow := gotorow - 1;
        end else
        begin
         LogItStamp('The month is NOT a match - Staying put',0);
        end;
end;

Procedure RunGotoIfHostAlive;
var
 gotorow : Integer;
 DoThisPing : String;
begin
 if not trystrtoint(ExcelParam02, gotorow) then
        begin
         LogItStamp('ERROR: Script parameter does not contain a number for goto value in row ' + inttostr(CurrentRow),0);
         exit
        end;
 if gotorow < 2 then
        begin
         LogItStamp('ERROR: Cannot goto any line before script start. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
if not IsValidIPv4Address(vartostr(ExcelParam01)) then
         begin
          if GetHostIP(vartostr(ExcelParam01)) = '' then
            begin
              LogItStamp('ERROR: Not a valid IP address and could not resolve hostname. Row ' + inttostr(CurrentRow),0);
              exit
            end;
         end;
 GotReplyFromPing := false;
 PingHost := VarToStr(ExcelParam01);
 DoThisPing := GetHostIP(PingHost);
 LogItStamp('Checking if host ' + '(' + pinghost + ') - ' + DoThisPing + ' is alive',0);
 PingHost := DoThisPing;
 frmMain.Ping.PingHost(PingHost);
 Chill(ChillTime);
     if not GotReplyFromPing then
      begin
      LogItStamp('Host NOT alive. Staying put',0);
      exit;
      end;
if GotReplyFromPing then LogItStamp('Host IS alive. Jumping to row ' + inttostr(gotorow),0);
   CurrentRow := gotorow - 1;
end;

Procedure RunGetPublicIP;
begin
  frmmain.ValueListEditorVariables.InsertRow('#ApplicationPublicIP#', GetPublicIPAddress, true);
  LogItStamp('Application Public IP in use is : ' + GetPublicIPAddress + '. Variable updated',0);
end;

Procedure RunDataSourceSetIndex;
var
 index : Integer;
begin
 if not trystrtoint(ExcelParam01, index) then
        begin
         LogItStamp('ERROR: The index is not a number. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if index > frmMain.DropDownData.Items.Count then
        begin
         LogItStamp('ERROR: The index chosen is too high. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
 if index < 1 then
        begin
         LogItStamp('ERROR: The index chosen is too low. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
frmMain.DropDownData.ItemIndex := index - 1;
LogItStamp('Changing data source to ' + frmMain.DropDownData.text,0);
updatevariables;
end;

Procedure RunGotoIfSSHConnected;
 var
 gotorow : integer;
begin
 if not trystrtoint(ExcelParam01, gotorow) then
        begin
         LogItStamp('ERROR: Script parameter does not contain a number for goto value in row ' + inttostr(CurrentRow),0);
         exit
        end;
 if gotorow < 2 then
        begin
         LogItStamp('ERROR: Cannot goto any line before script start. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
  if not frmMain.sshclient01.Connected then
        begin
         LogItStamp('SSH NOT connected. Staying put',0);
        end else begin
         LogItStamp('SSH connected. Jumping to row ' + inttostr(gotorow),0);
         CurrentRow := gotorow - 1;
        end;
 end;

Procedure RunInfoMessage;
var
 howhigh : integer;
begin
if frmMain.GroupBoxInformation.Visible = false then
  begin
   LogItStamp('ERROR: Information banner not visible. Skipping command at row ' + inttostr(CurrentRow),0);
   Exit
  end;
 frmMain.LabelInformation.Caption := vartostr(ExcelParam01);
   if fileexists(extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\data\' + vartostr(excelparam02)) then
      begin
        LogItStamp('Showing information "' + vartostr(ExcelParam01) + '" with picture "' + vartostr(excelparam02 + '"'),0);
        frmmain.ImageInformation.Picture.LoadFromFile(extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\data\' + vartostr(excelparam02));
        end else begin
                 LogItStamp('WARNING: Showing information "' + vartostr(ExcelParam01) + '", but picture "' + vartostr(excelparam02) + '" was not found',0);
                 end;
 if not trystrtoint(ExcelParam03, howhigh) then
        begin
          LogItStamp('ERROR: Pixel height reference not a number or missing. Skipping command at row ' + inttostr(CurrentRow),0);
          exit;
        end else begin
                  if howhigh > frmMain.CardLog.Height then howhigh := frmMain.CardLog.Height - 300;
                 end;
 frmMain.GroupBoxInformation.Height := howhigh;
end;

Procedure RunTelnetConnect;
begin
  ConnectToTelnetHost(vartostr(ExcelParam01));
end;

Procedure RunTelnetDisConnect;
begin
  DisconnectTelnetHost;
end;

Procedure RunTelnetCommand;
begin
verbose := false;
  if not frmMain.telnetclient01.IsConnected then begin
   LogItStamp('ERROR: Skipping command at row ' + inttostr(CurrentRow) + '. Telnet not connected',0);
  end else begin
   frmMain.telnetclient01.SendStr(vartostr(ExcelParam01) + #13#10);
   LogItStamp('Sending Telnet command "' + vartostr(ExcelParam01) + '" to connected host',0);
   chill(ChillTime);
  end;
end;

Procedure RunTelnetCommandVerbose;
begin
  if not frmMain.telnetclient01.IsConnected then begin
   LogItStamp('ERROR: Skipping command at row ' + inttostr(CurrentRow) + '. Telnet not connected',0);
  end else begin
   Verbose := true;
   LastVerboseStart := frmMain.MemoLog.Lines.Count - 1;
   LogItStamp('Sending Telnet Verbose command "' + vartostr(ExcelParam01) + '" to connected host',0);
   LogItStamp('Displaying data returned from host:',1);
   sshinputstring := '';
   frmMain.telnetclient01.SendStr(vartostr(ExcelParam01) + #13#10);
   chill(ChillTime);
  end;
Verbose := false;
LogIt('',0);
end;

Procedure RunGotoIfTelnetConnected;
 var
 gotorow : integer;
begin
 if not trystrtoint(ExcelParam01, gotorow) then
        begin
         LogItStamp('ERROR: Script parameter does not contain a number for goto value in row ' + inttostr(CurrentRow),0);
         exit
        end;
 if gotorow < 2 then
        begin
         LogItStamp('ERROR: Cannot goto any line before script start. Skipping command at row ' + inttostr(CurrentRow),0);
         exit;
        end;
  if not frmMain.telnetclient01.isConnected then
        begin
         LogItStamp('Telnet NOT connected. Staying put',0);
        end else begin
         LogItStamp('Telnet connected. Jumping to row ' + inttostr(gotorow),0);
         CurrentRow := gotorow - 1;
        end;
 end;

Procedure RunWaitForConnectionTelnet;
var
 AreWeAtMax : Integer;
begin
 AreWeAtMax := 0;
   if frmMain.telnetclient01.isConnected then LogItStamp('ERROR: Telnet connection to another host already active. Ignoring command at row ' + inttostr(CurrentRow),0);
   if frmMain.telnetclient01.isConnected then exit;
   LogItStamp('Trying Telnet connection to host ' + ExcelParam01 + ' with loop interval of ' + inttostr(timeout) + ' seconds',0);
   LogItStamp('Click the Stop button to abort',0);
   if MaxTimeout > 0 then LogItStamp('Max retry time set to ' + inttostr(MaxTimeout) + ' seconds',0);
         if MaxTimeout <= Timeout then begin
           LogItStamp('ERROR: Max retry time set lower than the connection time in settings. Adjusting to 180 seconds',0);
           MaxTimeout := 180;
         end;
    while not frmMain.telnetclient01.isConnected do
    begin
       if StopPressed then exit;
         frmMain.TimerSSH.Enabled := true;
         frmMain.telnetclient01.Host := VarToStr(ExcelParam01);
         frmMain.telnetclient01.Connect;
         if not frmMain.telnetclient01.isconnected then LogItStamp('Telnet host not ready yet. Retrying in ' + inttostr(timeout) + ' seconds',0);
         frmMain.TimerSSH.Enabled := false;
           if frmMain.telnetclient01.isConnected then LogItStamp('Telnet connected to host ' + ExcelParam01,0);
           if frmMain.telnetclient01.isConnected then Chill(ChillTime);
           Inc(AreWeAtMax, Timeout);
             if AreWeAtMax >= MaxTimeout then
               begin
                LogItStamp('Will not try again. Max retry time reached. Script will continue',0);
                exit;
               end;
       if not frmMain.telnetclient01.isConnected then Chill(Timeout);
    end;
end;

Procedure RunTelnetContent;
var
 xls : TXlsFile;
 Counter : integer;
 FileName : string;
 WaitTimeOut : integer;
 PushMemoLine : Integer;
 ChillCounter : Integer;
 NoFeedbackCounter : Integer;
 DelayedCounter : Integer;
 Delayed : Boolean;
begin
if not trystrtoint(ExcelParam02, WaitTimeOut) then WaitTimeOut := 300;
if ExcelParam01 = '' then begin
  LogItStamp('ERROR: Sending telnet content aborted. No content sheet specified',0);
  exit
end;
  LogItStamp('Getting Telnet content ready from content sheet ' + vartostr(ExcelParam01),0);
  xls := TXlsFile.Create;
  DelayedCounter := 0;
  NoFeedbackCounter := 0;
  Delayed := false;
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.xlsx';
  CopyFileToTemp(filename);
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.tmp';
      if SheetExists(filename,vartostr(ExcelParam01)) then begin
          xls.open(filename);
          xls.ActiveSheetByName := vartostr(ExcelParam01);
          Counter := 1;
          frmMain.MemoContent.clear;
                while not xls.GetCellValue(counter,1).IsEmpty do begin
                  frmMain.MemoContent.lines.add(vartostr(xls.GetCellValue(counter,1)));
                  inc(counter);
                  if StopPressed then exit;
                  application.ProcessMessages;
                end;
               xls.free;
          for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do begin
            frmMain.MemoContent.text := StringReplace(frmMain.MemoContent.text, frmMain.ValueListEditorVariables.Strings.KeyNames[Counter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter], [rfReplaceAll]);
            if StopPressed then exit;
            application.ProcessMessages;
          end;
             if not frmMain.telnetclient01.isConnected then LogItStamp('ERROR: Skipping Telnet Content push. Telnet not connected',0);
              if frmMain.telnetclient01.isConnected then begin
               LogItStamp('Start pushing content to host from sheet ' + vartostr(ExcelParam01) + ' containing ' + inttostr(frmMain.MemoContent.Lines.Count) + ' lines. Line send delay is ' + inttostr(WaitTimeOut) + ' milliseconds',0);
               PushMemoLine := frmMain.MemoLog.lines.Count;
                 for counter := 0 to frmMain.MemoContent.Lines.Count -1 do begin
                 feedbackreceived := false;
                 frmMain.TelnetClient01.SendStr(frmMain.MemoContent.lines.Strings[Counter] + #13#10);
                 sleep(WaitTimeOut);
                 application.ProcessMessages;
                   for chillcounter := 1 to 10 do begin
                     if not feedbackreceived then chill(1);
                     if not feedbackreceived then delayed := true;
                   end;
                 if delayed then inc(DelayedCounter);
                 if not feedbackreceived then inc(NoFeedbackCounter);
                 frmMain.MemoLog.lines[PushMemoLine] := '                      Written ' + inttostr(counter + 1) + ' of ' + inttostr(frmMain.MemoContent.Lines.Count) + ' lines to SHH host - Delayed transactions : ' + inttostr(DelayedCounter) + ' - No feedback from host : ' + inttostr(NoFeedbackCounter);
                 delayed := false;
                 if StopPressed then exit;
                 end;
                LogItStamp('Telnet content push done.',0);
                DeleteFile(filename);
              end;
      end else begin
        LogItStamp('ERROR: Content sheet not found. No content to work with. Aborting.',0);
      end;
end;

Procedure RunSSHTextFile;
var
 Counter : integer;
 FileName : string;
 WaitTimeOut : integer;
 PushMemoLine : Integer;
 ChillCounter : Integer;
 NoFeedbackCounter : Integer;
 DelayedCounter : Integer;
 Delayed : Boolean;
begin
if not trystrtoint(ExcelParam02, WaitTimeOut) then WaitTimeOut := 300;
if ExcelParam01 = '' then begin
  LogItStamp('ERROR: Sending SSH Text File aborted. No filename specified',0);
  exit
end;
  LogItStamp('Getting SSH content ready from filename ' + vartostr(ExcelParam01),0);
  DelayedCounter := 0;
  NoFeedbackCounter := 0;
  Delayed := false;
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + vartostr(ExcelParam01);
   if FileExists(filename) then
      begin
        frmMain.MemoContent.clear;
        frmMain.MemoContent.lines.LoadFromFile(filename);
         for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do
           begin
            frmMain.MemoContent.text := StringReplace(frmMain.MemoContent.text, frmMain.ValueListEditorVariables.Strings.KeyNames[Counter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter], [rfReplaceAll]);
            if StopPressed then exit;
            application.ProcessMessages;
           end;
             if not frmMain.sshclient01.Connected then LogItStamp('ERROR: Skipping SSH Content push. SSH not connected',0);
              if frmMain.sshclient01.Connected then begin
               LogItStamp('Start pushing content to host from filename ' + vartostr(ExcelParam01) + ' containing ' + inttostr(frmMain.MemoContent.Lines.Count) + ' lines. Line send delay is ' + inttostr(WaitTimeOut) + ' milliseconds',0);
               PushMemoLine := frmMain.MemoLog.lines.Count;
                 for counter := 0 to frmMain.MemoContent.Lines.Count -1 do begin
                 feedbackreceived := false;
                 frmMain.sshshell01.WriteString(frmMain.MemoContent.lines.Strings[Counter] + #13#10);
                 sleep(WaitTimeOut);
                 application.ProcessMessages;
                   for chillcounter := 1 to 10 do begin
                     if not feedbackreceived then chill(1);
                     if not feedbackreceived then delayed := true;
                   end;
                 if delayed then inc(DelayedCounter);
                 if not feedbackreceived then inc(NoFeedbackCounter);
                 frmMain.MemoLog.lines[PushMemoLine] := '                      Written ' + inttostr(counter + 1) + ' of ' + inttostr(frmMain.MemoContent.Lines.Count) + ' lines to SHH host - Delayed transactions : ' + inttostr(DelayedCounter) + ' - No feedback from host : ' + inttostr(NoFeedbackCounter);
                 delayed := false;
                 if StopPressed then exit;
                 end;
                LogItStamp('SSH content push done.',0);
              end;
      end else
        begin
         LogItStamp('ERROR: File not found. No content to work with. Aborting.',0);
        end;
end;

Procedure RunTelnetTextFile;
var
 Counter : integer;
 FileName : string;
 WaitTimeOut : integer;
 PushMemoLine : Integer;
 ChillCounter : Integer;
 NoFeedbackCounter : Integer;
 DelayedCounter : Integer;
 Delayed : Boolean;
begin
if not trystrtoint(ExcelParam02, WaitTimeOut) then WaitTimeOut := 300;
if ExcelParam01 = '' then begin
  LogItStamp('ERROR: Sending Telnet Text File aborted. No filename specified',0);
  exit
end;
  LogItStamp('Getting Telnet content ready from filename ' + vartostr(ExcelParam01),0);
  DelayedCounter := 0;
  NoFeedbackCounter := 0;
  Delayed := false;
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + vartostr(ExcelParam01);
   if FileExists(filename) then
      begin
        frmMain.MemoContent.clear;
        frmMain.MemoContent.lines.LoadFromFile(filename);
         for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do
           begin
            frmMain.MemoContent.text := StringReplace(frmMain.MemoContent.text, frmMain.ValueListEditorVariables.Strings.KeyNames[Counter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter], [rfReplaceAll]);
            if StopPressed then exit;
            application.ProcessMessages;
           end;
             if not frmMain.telnetclient01.isConnected then LogItStamp('ERROR: Skipping Telnet Content push. Telnet not connected',0);
              if frmMain.telnetclient01.isConnected then begin
               LogItStamp('Start pushing content to host from filename ' + vartostr(ExcelParam01) + ' containing ' + inttostr(frmMain.MemoContent.Lines.Count) + ' lines. Line send delay is ' + inttostr(WaitTimeOut) + ' milliseconds',0);
               PushMemoLine := frmMain.MemoLog.lines.Count;
                 for counter := 0 to frmMain.MemoContent.Lines.Count -1 do begin
                 feedbackreceived := false;
                 frmMain.telnetclient01.SendStr(frmMain.MemoContent.lines.Strings[Counter] + #13#10);
                 sleep(WaitTimeOut);
                 application.ProcessMessages;
                   for chillcounter := 1 to 10 do begin
                     if not feedbackreceived then chill(1);
                     if not feedbackreceived then delayed := true;
                   end;
                 if delayed then inc(DelayedCounter);
                 if not feedbackreceived then inc(NoFeedbackCounter);
                 frmMain.MemoLog.lines[PushMemoLine] := '                      Written ' + inttostr(counter + 1) + ' of ' + inttostr(frmMain.MemoContent.Lines.Count) + ' lines to SHH host - Delayed transactions : ' + inttostr(DelayedCounter) + ' - No feedback from host : ' + inttostr(NoFeedbackCounter);
                 delayed := false;
                 if StopPressed then exit;
                 end;
                LogItStamp('Telnet content push done.',0);
              end;
      end else
        begin
         LogItStamp('ERROR: File not found. No content to work with. Aborting.',0);
        end;
end;

Procedure RunGotoIFURL;
var
 RowNumber : Integer;
begin
 if not trystrtoint(ExcelParam02, RowNumber) then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The cell is either empty or not a number',0);
    end else begin
              if Rownumber < 2 then
                begin
                LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
                LogItStamp('The cell is a number lower than 2 and that reference is before script start',0);
                exit;
                end;
                  if IsWebsiteUp(vartostr(ExcelParam01)) then
                    begin
                     CurrentRow := RowNumber - 1;
                     LogItStamp('URL is responsive, jumping to row ' + inttostr(RowNumber),0);
                    end else
                        begin
                         LogItStamp('URL is NOT responsive, staying put',0);
                        end;
             end;
end;

Procedure RunSendSyslog;
var
 host : string;
 Severity : Integer;
 SyslogMessage : String;
begin
 if not trystrtoint(ExcelParam02, Severity) then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The cell is either empty or not a number',0);
    exit;
    end;
 if severity < 0 then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The severity can not be lower than 0',0);
    exit;
   end;
 if severity > 7 then
   begin
    LogItStamp('ERROR: An error was found for the number value setting at row ' + inttostr(CurrentRow),0);
    LogItStamp('The severity can not be higher than 7',0);
    exit;
   end;
 if not IsValidIPv4Address(vartostr(ExcelParam01)) then
         begin
          if GetHostIP(vartostr(ExcelParam01)) = '' then
            begin
              LogItStamp('ERROR: Not a valid IP address and could not resolve hostname. Row ' + inttostr(CurrentRow),0);
              exit
            end;
         end;
Host := vartostr(ExcelParam01);
SysLogMessage := vartostr(ExcelParam03);
LogItStamp('Sending SysLog message to host ' + host + ' with message "' + SysLogMessage + '"',0);
frmMain.Syslog.remotehost := host;
frmMain.Syslog.Active := true;
frmMain.SysLog.SendPacket(14, Severity, SysLogMessage);
frmMain.Syslog.Active := false;
end;

Procedure RunCMDCommand;
//can use & for nesting commands
begin
 LogItStamp('Sending command ' + vartostr(ExcelParam01) + ' to CMD',0);
 if ExecuteCommand(vartostr(ExcelParam01), frmMain.MemoLog, false) then
  begin
   chill(1);
   exit;
  end;
LogItStamp('ERROR: Could not send CMD command ' + vartostr(ExcelParam01),0);
end;

Procedure RunCMDCommandVerbose;
//can use & for nesting commands
begin
 LogItStamp('Sending command ' + vartostr(ExcelParam01) + ' to CMD and logging feedback',0);
 LastVerboseStart := frmMain.MemoLog.Lines.Count - 1;
 if ExecuteCommand(vartostr(ExcelParam01), frmMain.MemoLog, true) then
  begin
   chill(1);
   exit;
  end;
LogItStamp('ERROR: Could not send CMD command ' + vartostr(ExcelParam01),0);
end;

Procedure RunPSCommand;
//can use ; for nesting commands
begin
 LogItStamp('Sending command ' + vartostr(ExcelParam01) + ' to PowerShell',0);
 if ExecutePowerShellCommand(vartostr(ExcelParam01), frmMain.MemoLog, false) then
  begin
   chill(1);
   exit;
  end;
LogItStamp('ERROR: Could not send CMD command ' + vartostr(ExcelParam01),0);
end;

Procedure RunPSCommandVerbose;
//can use ; for nesting commands
begin
 LastVerboseStart := frmMain.MemoLog.Lines.Count - 1;
 LogItStamp('Sending command ' + vartostr(ExcelParam01) + ' to PowerShell and logging feedback',0);
 if ExecutePowerShellCommand(vartostr(ExcelParam01), frmMain.MemoLog, true) then
  begin
   chill(1);
   exit;
  end;
LogItStamp('ERROR: Could not send CMD command ' + vartostr(ExcelParam01),0);
end;

Procedure RunCMDContent;
var
 xls : TXlsFile;
 Counter : integer;
 FileName : string;
begin
if ExcelParam01 = '' then begin
  LogItStamp('ERROR: Sending CMD content aborted. No content sheet specified',0);
  exit
end;
  LogItStamp('Getting CMD content ready from content sheet ' + vartostr(ExcelParam01),0);
  xls := TXlsFile.Create;
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.xlsx';
  CopyFileToTemp(filename);
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.tmp';
     if SheetExists(filename,vartostr(ExcelParam01)) then
      begin
          xls.open(filename);
          xls.ActiveSheetByName := vartostr(ExcelParam01);
          Counter := 1;
          frmMain.MemoContent.clear;
                while not xls.GetCellValue(counter,1).IsEmpty do
                  begin
                  frmMain.MemoContent.lines.add(vartostr(xls.GetCellValue(counter,1)));
                  inc(counter);
                  if StopPressed then exit;
                  application.ProcessMessages;
                  end;
               xls.free;
               DeleteFile(filename);
          for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do
           begin
            frmMain.MemoContent.text := StringReplace(frmMain.MemoContent.text, frmMain.ValueListEditorVariables.Strings.KeyNames[Counter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter], [rfReplaceAll]);
            if StopPressed then exit;
            application.ProcessMessages;
           end;
            filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + 'tempcmd.bat';
            frmMain.MemoContent.lines.SaveToFile(filename);
            LogItStamp('Starting external CMD process to execute commands',0);
            ExternalProcess := ShellExecuteAndWait(filename);
            DeleteFile(filename);
            LogItStamp('External CMD process done.',0);
      end else
        begin
          LogItStamp('ERROR: Content sheet not found. No content to work with. Aborting.',0);
        end;
end;

Procedure RunCMDTextFile;
var
 Counter : integer;
 FileName : string;
begin
if ExcelParam01 = '' then begin
  LogItStamp('ERROR: No CMD Batch File specified. Just specify filename and place file in Data folder of the project.',0);
  exit
end;
  LogItStamp('Getting CMD Batch File ready. Stringreplace variables, if present.',0);
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + vartostr(ExcelParam01);
     if fileexists(Filename) then
        begin
         frmMain.MemoContent.Lines.LoadFromFile(Filename);
          for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do
           begin
            frmMain.MemoContent.text := StringReplace(frmMain.MemoContent.text, frmMain.ValueListEditorVariables.Strings.KeyNames[Counter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter], [rfReplaceAll]);
            if StopPressed then exit;
            application.ProcessMessages;
           end;
            filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\tempcmd.bat';
            frmMain.MemoContent.lines.SaveToFile(filename);
            LogItStamp('Starting external CMD Batch process to execute commands',0);
            ExternalProcess := ShellExecuteAndWait(Filename);
            DeleteFile(filename);
            LogItStamp('External CMD Batch process done.',0);
        end else
             begin
              LogItStamp('ERROR: Filename specified does not exist. Just specify filename and place file in Data folder of the project.',0);
             end;
end;

Procedure RunPSContent;
var
 xls : TXlsFile;
 Counter : integer;
 FileName : string;
begin
if ExcelParam01 = '' then begin
  LogItStamp('ERROR: Sending PowerShell content aborted. No content sheet specified',0);
  exit
end;
  LogItStamp('Getting PowerShell content ready from content sheet ' + vartostr(ExcelParam01),0);
  xls := TXlsFile.Create;
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.xlsx';
  CopyFileToTemp(filename);
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.tmp';
     if SheetExists(filename,vartostr(ExcelParam01)) then
      begin
          xls.open(filename);
          xls.ActiveSheetByName := vartostr(ExcelParam01);
          Counter := 1;
          frmMain.MemoContent.clear;
                while not xls.GetCellValue(counter,1).IsEmpty do
                  begin
                  frmMain.MemoContent.lines.add(vartostr(xls.GetCellValue(counter,1)));
                  inc(counter);
                  if StopPressed then exit;
                  application.ProcessMessages;
                  end;
               xls.free;
               DeleteFile(filename);
          for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do
           begin
            frmMain.MemoContent.text := StringReplace(frmMain.MemoContent.text, frmMain.ValueListEditorVariables.Strings.KeyNames[Counter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter], [rfReplaceAll]);
            if StopPressed then exit;
            application.ProcessMessages;
           end;
            filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + 'tempps.ps1';
            frmMain.MemoContent.lines.SaveToFile(filename);
            LogItStamp('Starting external PowerShell process to execute commands',0);
            ExternalProcess := ShellExecuteAndWait('Powershell.exe', '-executionpolicy bypass -File "' + Filename + '"');
            DeleteFile(filename);
            LogItStamp('External PowerShell process done.',0);
      end else
        begin
          LogItStamp('ERROR: Content sheet not found. No content to work with. Aborting.',0);
        end;
end;

Procedure RunPSTextFile;
var
 Counter : integer;
 FileName : string;
begin
if ExcelParam01 = '' then begin
  LogItStamp('ERROR: No PowerShell file specified. Just specify filename and place file in Data folder of the project.',0);
  exit
end;
  LogItStamp('Getting PowerShell file ready. Stringreplace variables, if present.',0);
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + vartostr(ExcelParam01);
     if fileexists(Filename) then
        begin
         frmMain.MemoContent.Lines.LoadFromFile(Filename);
          for Counter := 0 to frmMain.ValueListEditorVariables.Strings.Count -1 do
           begin
            frmMain.MemoContent.text := StringReplace(frmMain.MemoContent.text, frmMain.ValueListEditorVariables.Strings.KeyNames[Counter], frmMain.ValueListEditorVariables.Strings.ValueFromIndex[Counter], [rfReplaceAll]);
            if StopPressed then exit;
            application.ProcessMessages;
           end;
            filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\tempps.ps1';
            frmMain.MemoContent.lines.SaveToFile(filename);
            LogItStamp('Starting external PowerShell process to execute commands',0);
            ExternalProcess := ShellExecuteAndWait('Powershell.exe', '-executionpolicy bypass -File "' + Filename + '"');
            DeleteFile(filename);
            LogItStamp('External PowerShell Batch process done.',0);
        end else
             begin
              LogItStamp('ERROR: Filename specified does not exist. Just specify filename and place file in Data folder of the project.',0);
             end;
end;

Procedure RunGotoSelector;
var
 xls : TXlsFile;
 Counter : integer;
 FileName : string;
 Rownumber : Integer;
begin
 frmMain.labelselector.caption := vartostr(ExcelParam02);
 frmMain.ButtonSelectorContinue.Caption := vartostr(ExcelParam03);
 if vartostr(ExcelParam02) = '' then
 begin
 frmMain.labelselector.Caption := 'Select from drop down menu';
 LogItStamp('WARNING: No caption set for message label. Setting default text at row ' + inttostr(currentrow),0);
 end;
 if vartostr(ExcelParam03) = '' then
 begin
 frmMain.ButtonSelectorContinue.Caption := 'Click to continue';
 LogItStamp('WARNING: No caption set for button. Setting default text at row ' + inttostr(currentrow),0);
 end;
 if vartostr(ExcelParam01) = '' then
  begin
   LogItStamp('ERROR: No content sheet for drop down menu specified at row ' + inttostr(currentrow),0);
   exit;
  end;
 LogItStamp('Showing message ' + vartostr(ExcelParam02) + ' and drop down menu with options from sheet ' + vartostr(ExcelParam01),0);
 LogItStamp('Waiting for user to select and click continue',0);
 xls := TXlsFile.Create;
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.xlsx';
  CopyFileToTemp(filename);
  filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Tasks\' + frmMain.DropDownTask.Text + '.tmp';
    if SheetExists(filename,vartostr(ExcelParam01)) then
      begin
       xls.open(filename);
       xls.ActiveSheetByName := vartostr(ExcelParam01);
       Counter := 1;
       frmMain.ComboBoxSelector.clear;
          while not xls.GetCellValue(counter,1).IsEmpty do
           begin
            frmMain.ComboBoxSelector.items.add(vartostr(xls.GetCellValue(counter,1)));
            inc(counter);
            if StopPressed then exit;
            application.ProcessMessages;
           end;
       xls.free;
       DeleteFile(filename);
       if frmMain.comboboxSelector.items.count <> 0 then frmMain.comboboxSelector.itemindex := 0;
       frmMain.CardPanelRight.ActiveCard := frmmain.CardSelector;
      end else
        begin
         LogItStamp('ERROR: Content sheet not found. No data to fill drop down menu at row ' + inttostr(currentrow),0);
         exit;
        end;
if frmMain.ValueListEditorVariables.Values[frmMain.comboboxSelector.Text] = '' then
 begin
  LogItStamp('ERROR: LABELS not defined in script for the drop down selections at row ' + inttostr(currentrow),0);
  showmessage('One or more LABELS for the dropdown choices are not defined in the script.' + #13#10 + 'Script will continue, but maybe not with expected outcome');
  exit;
 end;
WaitForContinue;
frmMain.CardPanelRight.ActiveCard := frmMain.CardLog;
rownumber := strtoint(frmMain.ValueListEditorVariables.Values[frmMain.comboboxSelector.Text]);
currentrow := rownumber;
if not stoppressed then LogItStamp('User clicked to continue script. Jumping to row ' + inttostr(currentrow),0);
end;

Procedure RunLoadVariables;
var
Filename : String;
begin
filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + vartostr(ExcelParam01);
LoadVariablesFromFile(Filename);
LogItStamp('New variables added from file : ' + filename, 0);
end;

Procedure RunSaveVariables;
var
Filename : String;
begin
filename := extractfilepath(application.ExeName) + 'Projects\' + frmMain.DropDownProject.Text + '\Data\' + vartostr(ExcelParam01);
frmMain.ValueListEditorVariables.Strings.SaveToFile(Filename);
LogItStamp('All variables saved to file : ' + filename, 0);
end;

Procedure RunSSHConnectEx;
begin
  // SSHConnectEx(vartostr(ExcelParam01), vartostr(ExcelParam02), vartostr(ExcelParam03));
end;

Procedure RunSSHCommandEx;
begin
  // SSHCommandEx(vartostr(ExcelParam01), vartostr(ExcelParam02));
end;

Procedure RunSSHDisconnectEx;
begin
  // SSHDisconnectEx(vartostr(ExcelParam01));
end;

end.
