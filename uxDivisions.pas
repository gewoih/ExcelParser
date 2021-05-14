unit uxDivisions;

interface

procedure DrawDivisions;

var
	Divisions:	array of AnsiString;

implementation

uses MainForm, uxSQL, uxSuppliers, SysUtils, uxADO;

procedure DrawDivisions;
var
	R:	OleVariant;
    k:	integer;
begin
    Form1.cbDivisions.Clear;
    SetLength(Divisions, 0);
    k := 0;
        
    R := fcon.Execute(Format(Form1.scLoadDivisions.Items.Text, [Suppliers[Form1.SuppliersTree.FocusedNode.Index].pid]));
    while not R.EOF do
    begin
        SetLength(Divisions, k+1);
        Divisions[k] := AsStr(R, 1);

        Form1.cbDivisions.items.add(AsStr(R, 2) + ' [' + Divisions[k] + ']');

        R.MoveNext;
        Inc(k);
    end;
end;

end.
