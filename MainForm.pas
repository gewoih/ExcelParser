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
    	grid.colWidths[i] := max + grid.gridLineWidth + 30;
    end;
end;

procedure Excel_Open(file_name: string; grid: TStringGrid);
const xlCellTypeLastCell = $0000000B;
var
	Excel, Sheet:		   	OleVariant;
    i, j, grid_cursor_i:	integer;
begin
    Excel := CreateOleObject('Excel.Application');
    grid_cursor_i := 0;
    Grid.RowCount := 0;
    Grid.ColCount := 0;

    Excel.Workbooks.Open(file_name);

    for var sheet_number: integer := 1 to Excel.Workbooks[ExtractFileName(file_name)].Sheets.Count do
    begin
        Sheet := Excel.Workbooks[ExtractFileName(file_name)].WorkSheets[sheet_number];
        Sheet.Activate;
        Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;

        Grid.RowCount := Grid.RowCount + Excel.ActiveCell.Row;
        if Excel.ActiveCell.Column >= Grid.ColCount then
			Grid.ColCount := Excel.ActiveCell.Column;

        for i := 0 to Excel.ActiveCell.Row - 1 do
        begin
        	for j := 0 to Excel.ActiveCell.Column - 1 do
        		Grid.Cells[j, grid_cursor_i] := Sheet.Cells[i+1, j+1];
            Inc(grid_cursor_i);
        end;
    end;

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
