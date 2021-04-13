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
    procedure miAddSupplierClick(Sender: TObject);
  private
	procedure DrawPreview(Rows, Columns: integer);
  public

  end;

var
  Form1:				TForm1;
  fcon, PreviewArray:	OleVariant;

implementation

{$R *.dfm}

uses
    uxADO_cutted, uxSQL, ufAddSupplier;

procedure TForm1.DrawPreview(Rows, Columns: integer);
procedure SetTreeColumns;
var
	Column:	TVirtualTreeColumn;
begin
	try
        for var i: integer := 0 to Columns-1 do
        begin
            Column := PreviewTree.Header.Columns.Add;
            Column.Alignment := taLeftJustify;
            Column.Text := (i+1).ToString;
        end;
    finally
        PreviewTree.Invalidate;
    end;
end;

var
	N: PVirtualNode;
begin
    try
        PreviewTree.BeginUpdate;
        SetTreeColumns;

        PreviewTree.RootNodeCount := Rows;

        N := PreviewTree.GetFirst;
        while Assigned(N) do
        begin
            N.SetData(pointer(N.Index));
            N := N.NextSibling;
        end
    finally
        PreviewTree.Header.AutoFitColumns(false);
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
        if OpenDialog1.Execute then
        begin
        	total_time := Now();

            Excel := CreateOleObject('Excel.Application');
            Excel.Workbooks.Open(OpenDialog1.FileName, 0, true);
            Sheet := Excel.ActiveWorkbook.ActiveSheet;

            Rows := 50;
            //Rows := Sheet.UsedRange.Rows.Count;
            Columns := Sheet.UsedRange.Columns.Count;

            PreviewArray := Sheet.UsedRange.Value;

            DrawPreview(Rows, Columns);
        end;
    finally
    	ShowMessage(FormatDateTime('hh:mm:ss', Now() - total_time));
    	Excel.Quit;
    	Excel := Unassigned;
    end;
end;


procedure TForm1.FormCreate(Sender: TObject);
begin
    ConnectSQL(fcon);
end;

procedure TForm1.miAddSupplierClick(Sender: TObject);
begin
	try
        with TForm2.Create(nil) do
        begin
            if ShowModal = mrOk then
            begin
                fcon.Execute('insert into VTK_EXCEL.dbo.Suppliers values('
                + SuppliersTree.Text[SuppliersTree.FocusedNode, 0] + ')');
            end;
        end;
    finally

    end;
end;

procedure TForm1.PreviewTreeGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: string);
begin
    CellText := PreviewArray[Node.RowIndex+1, Column+1];
end;

end.
