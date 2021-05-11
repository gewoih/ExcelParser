unit uxPreview;

interface

procedure LoadPreview;
procedure DrawPreview;

var
	PreviewArray:	array of array[0..4] of AnsiString;

implementation

uses MainForm, VirtualTrees, System.Variants, uxExcel, SysUtils, uxExcelLinks, Dialogs;

procedure LoadPreview;
var
	rows, cols, val2, K:	integer;
    val1:					Extended;
begin
    rows := VarArrayHighBound(ExcelArray, 1);
    cols := 5;
    K := 0;

    SetLength(PreviewArray, rows);
    if 	(Excel_links[0] > -1) and
        (Excel_links[1] > -1) and
        (Excel_links[2] > -1) and
        (Excel_links[3] > -1) then
    begin
        for var i: integer := 1 to rows-1 do
        begin
            if  (String(ExcelArray[i, Excel_links[0]+1]) <> '') and
                (String(ExcelArray[i, Excel_links[1]+1]) <> '') and
                (TryStrToFloat(String(ExcelArray[i, Excel_links[2]+1]), val1)) and
                (TryStrToInt(String(ExcelArray[i, Excel_links[3]+1]), val2)) then
            begin
                PreviewArray[K, 0] := String(ExcelArray[i, Excel_links[0]+1]);
                PreviewArray[K, 1] := String(ExcelArray[i, Excel_links[1]+1]);
                PreviewArray[K, 2] := String(ExcelArray[i, Excel_links[2]+1]);
                PreviewArray[K, 3] := String(ExcelArray[i, Excel_links[3]+1]);
                if Excel_links[4] > -1 then
                    PreviewArray[K, 4] := String(ExcelArray[i, Excel_links[4]+1]);

                Inc(K);
            end;
        end;
        SetLength(PreviewArray, K);
    end;
end;

procedure DrawPreview;
var
	N:						PVirtualNode;
    row, col, delimiter:	integer;
    V:					  	OleVariant;
begin
	try
    	if Length(PreviewArray) > 0 then
        begin
            Form1.PreviewTree.BeginUpdate;
            Form1.PreviewTree.Clear;

            Form1.PreviewTree.RootNodeCount := Length(PreviewArray);

            N := Form1.PreviewTree.GetFirst;
            while Assigned(N) do
            begin
                N.SetData(pointer(N.Index));
                N := N.NextSibling;
            end
        end;
    finally
        Form1.PreviewTree.Header.AutoFitColumns(false);
        Form1.PreviewTree.EndUpdate;
        Form1.PreviewTree.Invalidate;
    end;
end;

end.
