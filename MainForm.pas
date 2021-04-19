unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls, System.Win.ComObj,
  Vcl.Grids, VirtualTrees, TFlatPanelUnit, Vcl.Menus, uBase, uSysCtrls;

type
  TForm1 = class(TForm)
    MainPanel: TPanel;
    OpenDialog1: TOpenDialog;
    SuppliersTree: TVirtualStringTree;
    PopupMenu1: TPopupMenu;
    miAddSupplier: TMenuItem;
    PreviewTree: TVirtualStringTree;
    FlatPanel1: TFlatPanel;
    scLoadSuppliers: TStringContainer;
    miChangeTemplate: TMenuItem;
    miDeleteSupplier: TMenuItem;
    scLoadExcel: TStringContainer;
    VirtualStringTree1: TVirtualStringTree;
    procedure FormCreate(Sender: TObject);
    procedure PreviewTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure miAddSupplierClick(Sender: TObject);
    procedure LoadSuppliers;
    procedure SuppliersTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure miChangeTemplateClick(Sender: TObject);
    procedure SuppliersTreeBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
    procedure SuppliersTreeFocusChanged(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Column: TColumnIndex);
    procedure PreviewTreePaintText(Sender: TBaseVirtualTree;
      const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      TextType: TVSTTextType);
  private
	procedure DrawPreview;
    type
        tSupplier = record
            id:		integer;
            pid:    integer;
            tin:    AnsiString;
            name:   AnsiString;
        end;
  public
    Suppliers:  array of tSupplier;
  end;

var
  Form1:				TForm1;
  fcon, PreviewArray:	OleVariant;

implementation

{$R *.dfm}

uses
    uxADO_cutted, uxSQL, ufAddSupplier;

procedure TForm1.DrawPreview;
procedure SetTreeColumns;
var
	Column:	TVirtualTreeColumn;
begin
	try
        PreviewTree.Header.Columns.Clear;
        for var i: integer := 0 to VarArrayHighBound(PreviewArray, 2)-1 do
        begin
            Column := PreviewTree.Header.Columns.Add;
            Column.Alignment := taLeftJustify;
            Column.CaptionAlignment := taCenter;
            Column.Text := Char(i+65);
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
        PreviewTree.Clear;
        SetTreeColumns;

        PreviewTree.RootNodeCount := VarArrayHighBound(PreviewArray, 1);

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

procedure TForm1.FormCreate(Sender: TObject);
begin
    ConnectSQL(fcon);
    LoadSuppliers;
end;

procedure TForm1.LoadSuppliers;
var
    R:	OleVariant;
    K:  integer;
begin
    try
        R := fcon.Execute(scLoadSuppliers.Items.Text);

        K := 0;
        SuppliersTree.BeginUpdate;
        SuppliersTree.Clear;
        while not R.EOF do
        begin
            SetLength(Suppliers, K+1);
            Suppliers[K].id := AsInt(R, 'id');
            Suppliers[K].pid := AsInt(R, 'pid');
            Suppliers[K].tin := AsStr(R, 'tin');
            Suppliers[K].name := AsStr(R, 'name');

            SuppliersTree.AddChild(nil, pointer(K));
            R.MoveNext;
            Inc(K);
        end;
    finally
        SuppliersTree.EndUpdate;
        SuppliersTree.Invalidate;
    end;
end;

procedure TForm1.miAddSupplierClick(Sender: TObject);
begin
	try
        with TForm2.Create(nil) do
        begin
            if ShowModal = mrOk then
            begin
                fcon.Execute('insert into Suppliers values('
                + SuppliersTree.Text[SuppliersTree.FocusedNode, 0] + ')');

                LoadSuppliers;
            end;
        end;
    finally

    end;
end;

procedure TForm1.miChangeTemplateClick(Sender: TObject);
var
	Excel, Sheet:					OleVariant;
    Rows, Columns, id:				integer;
    List:                           TStringList;
    S:                              String;
begin
	try
        if OpenDialog1.Execute and Assigned(SuppliersTree.FocusedNode) then
        begin
            Excel := CreateOleObject('Excel.Application');
            Excel.Workbooks.Open(OpenDialog1.FileName, 0, true);
            Sheet := Excel.ActiveWorkbook.ActiveSheet;
            List := TStringList.Create;

            Rows := 50;
            //Rows := Sheet.UsedRange.Rows.Count;
            Columns := Sheet.UsedRange.Columns.Count;

            PreviewArray := Sheet.UsedRange.Value;

            id := Suppliers[SuppliersTree.FocusedNode.Index].id;

            fcon.Execute('delete from Excel_templates where linkid = ' + id.ToString);

            List.Add('insert into Excel_templates values');
            for var i: integer := 1 to Rows do
            	for var j: integer := 1 to Columns do
                begin
                	if PreviewArray[i, j] = '' then continue;

                    if VarType(PreviewArray[i, j]) = varDouble then
                        S := QuotedStr(extended(PreviewArray[i, j]).ToString)
                    else
                        S := QuotedStr(PreviewArray[i, j]);

          			List.Add(Format('(' + id.ToString + ',%d, %d, %s),', [i, j, S]));
                end;
            List.Add('(' + id.ToString + '0, 1, null)');
            List.Add('(' + id.ToString + '0, 2, null)');
            List.Add('(' + id.ToString + '0, 3, null)');
            List.Add('(' + id.ToString + '0, 4, null)');

            //List.SaveToFile('test.sql');
            fcon.Execute(List.Text);

            DrawPreview;

            Excel.Quit;
    		Excel := Unassigned;
        end;
    finally
    end;
end;

procedure TForm1.PreviewTreeGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: string);
var
	V: Variant;
begin
    V := PreviewArray[Node.RowIndex+1, Column+1];

    if VarIsNull(V) then
    	CellText := ''
    else
//    	CellText := VarTypeAsText(VarType(V));
    	CellText := PreviewArray[Node.RowIndex+1, Column+1];
end;

procedure TForm1.PreviewTreePaintText(Sender: TBaseVirtualTree;
  const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  TextType: TVSTTextType);
begin
    if VarType(PreviewArray[Node.Index+1, Column+1]) = varDouble then
    begin
        TargetCanvas.Font.Color := clRed;
        //TargetCanvas.TextRect
    end;
end;

procedure TForm1.SuppliersTreeBeforeCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
begin
var
	R := CellRect;
    if Node.Index mod 2 = 0 then TargetCanvas.Brush.Color := $E0FFE0 else TargetCanvas.Brush.Color := $E0FFFF;
    if (Node=Sender.FocusedNode) or (Column=Sender.FocusedColumn) then TargetCanvas.Brush.Color := TargetCanvas.Brush.Color - $200020;
    if (Node=Sender.FocusedNode) and (Column=Sender.FocusedColumn) then TargetCanvas.Brush.Color := $FFFFFF;
    TargetCanvas.FillRect(R);
end;

procedure TForm1.SuppliersTreeFocusChanged(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex);
procedure LoadPreview;
var
	R:				OleVariant;
    rows, cols:     integer;
begin
    R := fcon.Execute(Format(scLoadExcel.Items.Text, [Suppliers[Node.Index].id]));
    rows := R.RecordCount;
    cols := R.Fields.Count;

    PreviewArray := VarArrayCreate([1, rows, 1, cols], varVariant);

    while not R.EOF do
    begin
        for var i := 0 to cols-1 do
        begin
            PreviewArray[R.AbsolutePosition, i+1] := R.Fields[i].Value;
        end;
        R.MoveNext;
    end;
end;

begin
    LoadPreview;
    DrawPreview;
end;

procedure TForm1.SuppliersTreeGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: string);
begin
    case Column of
    0: CellText := Suppliers[Node.Index].id.ToString;

    1: CellText := Suppliers[Node.Index].pid.ToString;

    2: CellText := Suppliers[Node.Index].tin;

    3: CellText := Suppliers[Node.Index].name;
    end;
end;

begin
    FormatSettings.DecimalSeparator := '.';

end.
