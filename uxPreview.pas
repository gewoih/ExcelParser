unit uxPreview;

interface

procedure LoadPreview;
procedure DrawPreview;
procedure UploadPrice;

var
	PreviewArray:	array of array[0..5] of AnsiString;

implementation

uses MainForm, VirtualTrees, System.Variants, uxExcel, SysUtils, uxExcelLinks, Dialogs, Classes,
uxUtils, uxDivisions, Vcl.Controls, uxSQL;

procedure LoadPreview;
function ParseYear(input_str: String): AnsiString;
var
	S:	AnsiString;
    C:	AnsiChar;
    k:	integer;
begin
	S := input_str;

    Result := '';

    for C in S do
    begin
        if byte(C) in [48..57] then
        	Result := Result + C
        else
        	Result := '';
        if Length(Result) = 4 then
        	break;
    end;

    if Length(Result) <> 4 then
    	Result := '';
end;
var
	rows, val2, K:	integer;
    val1:			Extended;
begin
    rows := VarArrayHighBound(ExcelArray, 1);
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
                (TryStrToFloat(String(ExcelArray[i, Excel_links[3]+1]), val1)) then
            begin
                PreviewArray[K, 0] := String(ExcelArray[i, Excel_links[0]+1]);
                PreviewArray[K, 1] := String(ExcelArray[i, Excel_links[1]+1]);
                PreviewArray[K, 2] := String(ExcelArray[i, Excel_links[2]+1]);
                PreviewArray[K, 3] := String(ExcelArray[i, Excel_links[3]+1]);
                if Excel_links[4] > -1 then
                	PreviewArray[K, 4] := String(ExcelArray[i, Excel_links[4]+1]);
                if Excel_links[5] > -1 then
                begin
                    PreviewArray[K, 5] := ParseYear(String(ExcelArray[i, Excel_links[5]+1]));
                    if PreviewArray[K, 5] <> '' then
                	begin
                        PreviewArray[K, 0] := PreviewArray[K, 0] + ', ' + PreviewArray[K, 5] + 'г.';
                        PreviewArray[K, 1] := PreviewArray[K, 1] + '-' + PreviewArray[K, 5];
                    end;
                end;

                Inc(K);
            end;
        end;
        SetLength(PreviewArray, K);
    end;
end;

procedure DrawPreview;
var
	N:	PVirtualNode;
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

procedure UploadPrice;
var
  S: 				string;
  So: 				tStringlist;
  regid:			AnsiString;
  buttonSelected:	Integer;
begin
	try
    	buttonSelected := MessageDlg('¬ы действительно хотите загрузить ' + Length(PreviewArray).ToString +
        ' позиций дл€ подразделени€ ' + Form1.cbDivisions.Text + '?', mtConfirmation, mbOKCancel, 0);

  		if buttonSelected = mrOK then
        begin
            So := tStringlist.Create;
            regid := Divisions[Form1.cbDivisions.ItemIndex];

            So.Add('if object_id(''tempdb..##temp_rest'') <> 0 drop table ##temp_rest');
            So.Add('select * into ##temp_rest from (values');

            for var i := 0 to High(PreviewArray) do
            begin
                So.Add('(' +
                QuotedStr(PreviewArray[i][1]) + ',' +
                '1' + ',' +
                PreviewArray[i][4] + ',' +
                PreviewArray[i][3] + ',' +
                PreviewArray[i][2] + ',' +
                QuotedStr(PreviewArray[i][0]) + ')' +
                iif(i = High(PreviewArray), '', ','));
            end;

            So.Add(') x (art, ap, qty, costin, cost, name)');
            So.Add('exec VTK.dbo.PreparePrice ' + QuotedStr(regid) + ',' + QuotedStr(FormatDateTime('yyyyMMdd', Now)));
            So.SaveToFile('D:\test\test.sql');
            fcon.Execute(So.Text);

            ShowMessage('ѕрайс успешно загружен.');
        end;
    finally
    	So.Free;
    end;
end;
end.
