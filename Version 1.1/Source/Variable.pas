unit Variable;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Grids, Vcl.ValEdit,
  Vcl.Menus, snippets;

type
  TfrmVariable = class(TForm)
    ButtonCloseVariableViewer: TButton;
    ValueListLiveVariables: TValueListEditor;
    PopupMenu1: TPopupMenu;
    PopMenuExportToNotepad: TMenuItem;
    procedure ButtonCloseVariableViewerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure PopMenuExportToNotepadClick(Sender: TObject);
  private
  public
  end;

var
  frmVariable: TfrmVariable;

procedure CopyValueListEditor(Source, Destination: TValueListEditor);

implementation

uses main;

{$R *.dfm}

procedure CopyValueListEditor(Source, Destination: TValueListEditor);
var
  i: Integer;
  key, value: string;
begin
  Destination.Strings.BeginUpdate;
  try
    Destination.Strings.Clear;
    for i := 1 to Source.Strings.Count do
    begin
      key := Source.Keys[i];
      value := Source.Values[key];
      Destination.Strings.AddPair(key, value);
//      application.ProcessMessages;
    end;
  finally
    Destination.Strings.EndUpdate;
  end;
end;

procedure TfrmVariable.ButtonCloseVariableViewerClick(Sender: TObject);
begin
 frmVariable.Visible := false;
end;

procedure TfrmVariable.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 frmVariable.Visible := false;
end;

procedure TfrmVariable.FormShow(Sender: TObject);
begin
 CopyValueListEditor(frmMain.ValueListEditorVariables, frmVariable.ValueListLiveVariables);
end;

procedure TfrmVariable.PopMenuExportToNotepadClick(Sender: TObject);
begin
OpenNotepadWithValueListEditor(frmVariable.ValueListLiveVariables);
end;

end.
