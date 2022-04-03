program TimeSheet;

uses
  Forms,
  MAIN in 'MAIN.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Time Sheet ';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
