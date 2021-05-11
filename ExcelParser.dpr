program ExcelParser;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {Form1},
  uxSQL in 'uxSQL.pas',
  ufAddSupplier in 'ufAddSupplier.pas' {Form2},
  uxExcelLinks in 'uxExcelLinks.pas',
  uxSuppliers in 'uxSuppliers.pas',
  uxExcel in 'uxExcel.pas',
  uxADO in 'uxADO.pas',
  uxPreview in 'uxPreview.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
