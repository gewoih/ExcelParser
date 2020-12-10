unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls, System.Win.ComObj,
  Vcl.Grids;

type
  TForm1 = class(TForm)
    MainPanel: TPanel;
    ChooseButton: TButton;
    StringGrid1: TStringGrid;
    OpenDialog1: TOpenDialog;
    procedure ChooseButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure SetMaxColumnWidth(grid: TStringGrid);
var
    i, j, temp, max:    integer;
begin
	for i := 0 to Grid.ColCount - 1 do
    begin
        max := 0;
        for j := 0 to Grid.RowCount - 1 do
        begin
            temp := Grid.Canvas.TextWidth(Grid.Cells[i, j]);
            if temp > max then
                max := temp;
        end;
    	grid.colWidths[i] := max + grid.gridLineWidth + 1;
    end;
end;

procedure Excel_Open(file_name: string; grid: TStringGrid);
const xlCellTypeLastCell = $0000000B;
var
	Excel, Sheet:							OleVariant;
    i, j, last_row, last_column, temp, max:	integer;
begin
    Excel := CreateOleObject('Excel.Application');

	Excel.Workbooks.Open(file_name, 0, True);

    Sheet := Excel.Workbooks[ExtractFileName(file_name)].WorkSheets[1];
    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;

    last_row := Excel.ActiveCell.Row;
    last_column := Excel.ActiveCell.Column;

    Grid.RowCount := last_row;
    Grid.ColCount := last_column;

    for i := 1 to last_row do
        for j := 1 to last_column do
            Grid.Cells[j-1, i-1] := Sheet.Cells[i, j];

    SetMaxColumnWidth(grid);

    Excel.Quit;
    Excel := Unassigned;
    Sheet := Unassigned;
end;

procedure TForm1.ChooseButtonClick(Sender: TObject);
begin
    if OpenDialog1.Execute then Excel_Open(OpenDialog1.FileName, StringGrid1);
end;
end.
