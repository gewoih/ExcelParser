program ExcelParser;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {Form1},
  uxADO_cutted in 'uxADO_cutted.pas',
  uxSQL in 'uxSQL.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
