unit uxSuppliers;

interface

procedure LoadSuppliers;
procedure AddSupplier;
procedure DeleteSupplier(id: integer);

type
    tSupplier = record
        id:		integer;
        pid:    integer;
        tin:    AnsiString;
        name:   AnsiString;
end;

var
	Suppliers: array of tSupplier;

implementation

uses MainForm, uxADO, Dialogs, SysUtils, ufAddSupplier, Vcl.Controls, uxSQL;

procedure AddSupplier;
begin
	try
        with TForm2.Create(nil) do
        begin
            if ShowModal = mrOk then
            begin
                fcon.Execute('insert into Suppliers values('
                + SuppliersTree.Text[SuppliersTree.FocusedNode, 0] + ')');

                LoadSuppliers; //Подгружать не всех, а только выбранного
            end;
        end;
    finally

    end;
end;

procedure DeleteSupplier(id: integer);
begin
	try
        fcon.Execute('delete from Suppliers where id = ' + id.ToString +
        '; delete from Excel_templates where linkid = ' + id.ToString);

        ShowMessage('Поставщик успешно удален!');

        LoadSuppliers;
    finally
    end;
end;

procedure LoadSuppliers;
var
    R:	OleVariant;
    K:  integer;
begin
    try
    	Form1.SuppliersTree.BeginUpdate;
        Form1.SuppliersTree.Clear;

       	R := fcon.Execute(Form1.scLoadSuppliers.Items.Text);

        K := 0;
        while not R.EOF do
        begin
            SetLength(Suppliers, K+1);
            Suppliers[K].id := AsInt(R, 'id');
            Suppliers[K].pid := AsInt(R, 'pid');
            Suppliers[K].tin := AsStr(R, 'tin');
            Suppliers[K].name := AsStr(R, 'name');

            Form1.SuppliersTree.AddChild(nil, pointer(K));

            R.MoveNext;
            Inc(K);
        end;
    finally
    	Form1.SuppliersTree.EndUpdate;
        Form1.SuppliersTree.Invalidate;
    end;
end;

end.
