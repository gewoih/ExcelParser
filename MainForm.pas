unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls, System.Win.ComObj,
  Vcl.Grids, VirtualTrees, TFlatPanelUnit, Vcl.Menus, uBase, uSysCtrls, ClipBrd, System.UITypes,
  Vcl.Tabs;

type
  TForm1 = class(TForm)
    MainPanel: TPanel;
    OpenDialog1: TOpenDialog;
    SuppliersTree: TVirtualStringTree;
    SuppliersTreePopupMenu: TPopupMenu;
    miAddSupplier: TMenuItem;
    ExcelTree: TVirtualStringTree;
    FlatPanel1: TFlatPanel;
    scLoadSuppliers: TStringContainer;
    miAddExcel: TMenuItem;
    miDeleteSupplier: TMenuItem;
    scLoadExcel: TStringContainer;
    PreviewTree: TVirtualStringTree;
    FlatPanel2: TFlatPanel;
    Label3: TLabel;
    btSaveLinks: TButton;
    FlatPanel3: TFlatPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    tName: TEdit;
    tArticle: TEdit;
    tBasePrice: TEdit;
    tQuantity: TEdit;
    btUploadPrice: TButton;
    Label5: TLabel;
    tYear: TEdit;
    scAddSupplier: TStringContainer;
    label6: TLabel;
    tPriceIn: TEdit;
    Label7: TLabel;
    cbDivisions: TComboBox;
    scLoadDivisions: TStringContainer;
    Splitter1: TSplitter;
    tVolume: TEdit;
    tStrength: TEdit;
    Label8: TLabel;
    Label9: TLabel;
    FlatPanel4: TFlatPanel;
    tbExcelTabs: TTabSet;
    procedure FormCreate(Sender: TObject);
    procedure ExcelTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure miAddSupplierClick(Sender: TObject);
    procedure SuppliersTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure miAddExcelClick(Sender: TObject);
    procedure SuppliersTreeBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
    procedure SuppliersTreeFocusChanged(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Column: TColumnIndex);
    procedure btSaveLinksClick(Sender: TObject);
    procedure PreviewTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure miDeleteSupplierClick(Sender: TObject);
    procedure ExcelTreeKeyPress(Sender: TObject; var Key: Char);
    procedure btUploadPriceClick(Sender: TObject);
    procedure tbExcelTabsChange(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
end;

var
	Form1:		TForm1;
    fLastNode:	PVirtualNode;
    tab_index:	integer;

implementation

uses ufAddSupplier, uxExcelLinks, uxSuppliers, uxExcel, uxPreview, uxADO, uxSQL, uxDivisions;

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
	LoadSuppliers
end;

procedure TForm1.miAddSupplierClick(Sender: TObject);
begin
	AddSupplier;
end;

procedure TForm1.PreviewTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
  Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
var
	V:	Variant;
begin
	V := PreviewArray[Node.RowIndex, Column];

    if (VarIsNull(V)) or (VarIsEmpty(V)) or (VarIsClear(V)) or (VarType(V) = varError) then
        CellText := ''
    else
        CellText := VarToStr(V);
end;

procedure TForm1.miAddExcelClick(Sender: TObject);
begin
    AddExcel;
end;

procedure TForm1.miDeleteSupplierClick(Sender: TObject);
begin
	if Assigned(SuppliersTree.FocusedNode) then
		DeleteSupplier(Suppliers[SuppliersTree.FocusedNode.Index].id);
end;

procedure TForm1.btUploadPriceClick(Sender: TObject);
begin
	if cbDivisions.ItemIndex <> -1 then
		UploadPrice
    else
    	ShowMessage('Выберите подразделение!');
end;

procedure TForm1.ExcelTreeGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: string);
var
	V:	Variant;
    k:	integer;
begin
    V := ExcelArray[tab_index][Node.RowIndex+1, Column+1];

    if (VarIsNull(V)) or (VarIsEmpty(V)) or (VarIsClear(V)) or (VarType(V) = varError) then
        CellText := ''
    else
        CellText := VarToStr(V);
end;

procedure TForm1.ExcelTreeKeyPress(Sender: TObject; var Key: Char);
var
    position:	AnsiString;
begin
    if Assigned(ExcelTree.FocusedNode) then
    begin
    	position := ExcelTree.ActiveColumn.Index.ToString;

    	case Ord(Key) of
            49: tName.Text := position;
            50: tArticle.Text := position;
            51: tBasePrice.Text := position;
            52: tPriceIn.Text := position;
            53:	tQuantity.Text := position;
            54: tYear.Text := position;
            55:	tVolume.Text := position;
            56:	tStrength.Text := position;
        end;
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

procedure TForm1.btSaveLinksClick(Sender: TObject);
begin
    SaveLinks;

    if ExcelTree.RootNodeCount <> 0 then
    begin
	    LoadPreview;
    	DrawPreview;
    end;
end;

procedure TForm1.SuppliersTreeFocusChanged(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex);
begin
    if not Assigned(fLastNode) or (fLastNode.Index <> Node.Index) then
    begin
    	tbExcelTabs.Tabs.Clear;
        ExcelTree.Clear;
        PreviewTree.Clear;
        DrawLinks;
        DrawDivisions;

        fLastNode := Node;
    end;
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

procedure TForm1.tbExcelTabsChange(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
begin
    if NewTab <> -1 then
    begin
    	tab_index := NewTab;

        DrawExcel;
    end;
end;

begin
    FormatSettings.DecimalSeparator := '.';
end.
