program ExcelParser;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {Form1},
  uxADO_cutted in 'uxADO_cutted.pas',
  uxSQL in 'uxSQL.pas',
  ufAddSupplier in 'ufAddSupplier.pas' {Form2},
  ufChangeLinks in 'ufChangeLinks.pas' {Form3};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.
