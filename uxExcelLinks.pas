unit uxExcelLinks;

interface
	procedure DrawLinks;
    procedure SaveLinks;

var
	Excel_links:	array [0..7] of integer;

implementation

uses MainForm, uxSuppliers, uxADO, SysUtils, System.Variants, Dialogs, VirtualTrees, uxExcel, uxPreview,
uxSQL;

procedure DrawLinks;
var
	R:		OleVariant;
    index:	integer;
begin
	index := Suppliers[Form1.SuppliersTree.FocusedNode.Index].id;
    fillchar(Excel_links, sizeof(Excel_links), -1);

    R := fcon.Execute('select * from Excel_templates where row = 0 and linkid = ' + index.ToString);

    for var i: integer := 0 to R.RecordCount-1 do
    begin
        Excel_links[i] := AsInt(R, 'val');
        R.MoveNext;
    end;

    Form1.tName.Text := Excel_Links[0].ToString;
    Form1.tArticle.Text := Excel_Links[1].ToString;
    Form1.tBasePrice.Text := Excel_Links[2].ToString;
    Form1.tPriceIn.Text := Excel_Links[3].ToString;
    Form1.tQuantity.Text := Excel_Links[4].ToString;
    Form1.tYear.Text := Excel_Links[5].ToString;
    Form1.tVolume.Text := Excel_links[6].ToString;
    Form1.tStrength.Text := Excel_links[7].ToString;
end;

procedure SaveLinks;
var
	id:	integer;
begin
	try
    	Excel_links[0] := StrToInt(Form1.tName.Text);
        Excel_links[1] := StrToInt(Form1.tArticle.Text);
        Excel_links[2] := StrToInt(Form1.tBasePrice.Text);
        Excel_links[3] := StrToInt(Form1.tPriceIn.Text);
        Excel_links[4] := StrToInt(Form1.tQuantity.Text);
        Excel_links[5] := StrToInt(Form1.tYear.Text);
        Excel_links[6] := StrToInt(Form1.tVolume.Text);
        Excel_links[7] := StrToInt(Form1.tStrength.Text);

    	if (Excel_links[0] > -1) and (Excel_links[1] > -1) and
       	   (Excel_links[2] > -1) and (Excel_links[3] > -1) then
        begin
            id := Suppliers[Form1.SuppliersTree.FocusedNode.Index].id;

            fcon.Execute('delete from Excel_templates where row = 0 and linkid = ' + id.ToString);

            fcon.Execute('insert into Excel_templates values'
            + '(' + id.ToString + ', 0, 1, ' + QuotedStr(Excel_links[0].ToString) + '),'
            + '(' + id.ToString + ', 0, 2, ' + QuotedStr(Excel_links[1].ToString) + '),'
            + '(' + id.ToString + ', 0, 3, ' + QuotedStr(Excel_links[2].ToString) + '),'
            + '(' + id.ToString + ', 0, 4, ' + QuotedStr(Excel_links[3].ToString) + '),'
            + '(' + id.ToString + ', 0, 5, ' + QuotedStr(Excel_links[4].ToString) + '),'
            + '(' + id.ToString + ', 0, 6, ' + QuotedStr(Excel_links[5].ToString) + '),'
            + '(' + id.ToString + ', 0, 7, ' + QuotedStr(Excel_links[6].ToString) + '),'
            + '(' + id.ToString + ', 0, 8, ' + QuotedStr(Excel_links[7].ToString) + ')');

            ShowMessage('Столбцы успешно привязаны.');
        end
        else
        	ShowMessage('Введите все ячейки для привязки!');
    finally
    end;
end;

end.
