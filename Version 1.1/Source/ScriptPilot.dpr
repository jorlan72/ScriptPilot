program ScriptPilot;

uses
  Vcl.Forms,
  main in 'main.pas' {FrmMain},
  snippets in 'snippets.pas',
  Vcl.Themes,
  Vcl.Styles,
  engine in 'engine.pas',
  adddata in 'adddata.pas',
  addtask in 'addtask.pas',
  Variable in 'Variable.pas' {frmVariable};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Windows10 BlackPearl');
  Application.Title := 'ScriptoMatic';
  Application.CreateForm(TFrmMain, FrmMain);
  Application.CreateForm(TfrmVariable, frmVariable);
  Application.Run;
end.
