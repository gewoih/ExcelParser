unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls, System.Win.ComObj,
  Vcl.Grids, VirtualTrees, TFlatPanelUnit, Vcl.Menus;

type
  TForm1 = class(TForm)
    MainPanel: TPanel;
    btExcelOpen: TButton;
    OpenDialog1: TOpenDialog;
    SuppliersTree: TVirtualStringTree;
    PopupMenu1: TPopupMenu;
    miAddSupplier: TMenuItem;
    PreviewTree: TVirtualStringTree;
    FlatPanel1: TFlatPanel;
    procedure btExcelOpenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure PreviewTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
  private
	procedure DrawPreview(Sheet: OleVariant; Rows, Columns: integer);
  public

  end;

var
  Form1: 				TForm1;
  fcon_local, fcon_vtk:	OleVariant;
  PreviewArray:			array of array of AnsiString;

implementation

{$R *.dfm}

uses
    uxADO_cutted, uxSQL;

procedure TForm1.DrawPreview(Sheet: OleVariant; Rows, Columns: integer);
procedure SetTreeColumns(Sheet: OleVariant; Columns: integer);
var
	Column:	TVirtualTreeColumn;
begin
	try
        for var i := 0 to Columns do
        begin
            Column := PreviewTree.Header.Columns.Add;
            Column.Tag := i;
            Column.Alignment := taLeftJustify;
            Column.Text := (i+1).ToString;
        end;
    finally
        PreviewTree.Invalidate;
    end;
end;

begin
    try
        PreviewTree.BeginUpdate;
        SetTreeColumns(Sheet, Columns);
        SetLength(PreviewArray, Rows);

        for var j := 0 to 100 do
        begin
        	SetLength(PreviewArray[j], Columns);
            for var i := 0 to Columns do
            begin
                PreviewArray[j, i] := Sheet.Cells[j+1, i+1];
            end;
            PreviewTree.AddChild(nil, pointer(j));
        end
    finally
        PreviewTree.Header.AutoFitColumns;
        PreviewTree.EndUpdate;
        PreviewTree.Invalidate;
    end;
end;

procedure TForm1.btExcelOpenClick(Sender: TObject);
var
	Excel, Sheet:					OleVariant;
    Rows, Columns:					integer;
    total_time:                     TDateTime;
begin
	try
        total_time := Now();
        if OpenDialog1.Execute then
        begin
            Excel := CreateOleObject('Excel.Application');
            Excel.Workbooks.Open(OpenDialog1.FileName, 0, true);
            Sheet := Excel.ActiveWorkbook.ActiveSheet;

            Rows := Sheet.UsedRange.Rows.Count;
            Columns := Sheet.UsedRange.Columns.Count;

            DrawPreview(Sheet, Rows, Columns);
        end;
        ShowMessage(FormatDateTime('hh:mm:ss', Now() - total_time));
    finally
    	Excel.Quit;
    	Excel := Unassigned;
    end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
    ConnectSQL(fcon_local, 'local');
end;

procedure TForm1.PreviewTreeGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: string);
begin
    CellText := PreviewArray[Node.RowIndex, Column];
end;

end.
