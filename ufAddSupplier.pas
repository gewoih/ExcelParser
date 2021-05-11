unit ufAddSupplier;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, VirtualTrees, uBase,
  uSysCtrls;

type
  TForm2 = class(TForm)
    SuppliersTree: TVirtualStringTree;
    btAddSupplier: TButton;
    scGetSuppliersList: TStringContainer;
    procedure btAddSupplierClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure DrawSuppliersTree;
    procedure SuppliersTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure SuppliersTreeBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
  private
    type
    	tSuppliers = Record
            id: 	integer;
            tin:    AnsiString;
            name:   AnsiString;
        End;
  public
    Suppliers:  array of tSuppliers;
  end;

var
  Form2:	TForm2;

implementation

{$R *.dfm}

uses uxADO, uxSQL;

procedure TForm2.btAddSupplierClick(Sender: TObject);
begin
    ModalResult := mrOK;
end;

procedure TForm2.DrawSuppliersTree;
var
    R:	OleVariant;
    K:  integer;
begin
	try
        K := 0;
        R := fcon.Execute(scGetSuppliersList.Items.Text);

        SuppliersTree.BeginUpdate;
        while not R.EOF do
        begin
            SetLength(Suppliers, K+1);

            Suppliers[K].id := AsInt(R, 0);
            Suppliers[K].tin := AsStr(R, 1);
            Suppliers[K].name := AsStr(R, 2);

            SuppliersTree.AddChild(nil, pointer(K));

            Inc(K);
            R.MoveNext;
        end;
    finally
        SuppliersTree.EndUpdate;
        SuppliersTree.Invalidate;
    end;
end;

procedure TForm2.FormCreate(Sender: TObject);
begin
	try
        DrawSuppliersTree;
    finally
    end;
end;
procedure TForm2.SuppliersTreeBeforeCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
begin
	var R := CellRect;
    if Node.Index mod 2 = 0 then TargetCanvas.Brush.Color := $E0FFE0 else TargetCanvas.Brush.Color := $E0FFFF;
    if (Node=Sender.FocusedNode) or (Column=Sender.FocusedColumn) then TargetCanvas.Brush.Color := TargetCanvas.Brush.Color - $200020;
    if (Node=Sender.FocusedNode) and (Column=Sender.FocusedColumn) then TargetCanvas.Brush.Color := $FFFFFF;
    TargetCanvas.FillRect(R);
end;

procedure TForm2.SuppliersTreeGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: string);
begin
    case Column of
    0: CellText := Suppliers[Node.Index].id.ToString;

    1: CellText := Suppliers[Node.Index].tin;

    2: CellText := Suppliers[Node.Index].name;
    end;
end;

end.
