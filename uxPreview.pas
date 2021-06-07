unit uxPreview;

interface

procedure LoadPreview;
procedure DrawPreview;
procedure UploadPrice;

var
	PreviewArray:	array of array[0..7] of AnsiString;

implementation

uses MainForm, VirtualTrees, System.Variants, uxExcel, SysUtils, uxExcelLinks, Dialogs, Classes,
uxUtils, uxDivisions, Vcl.Controls, uxSQL, System.UITypes;

function ParseYear(input_str: AnsiString): AnsiString;
var
    C:	AnsiChar;
begin
    Result := '';

    for C in input_str do
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

function GetMaxLink: integer;
begin
	Result := -1;
	for var i: integer := 0 to 3 do
	begin
		if Excel_links[i] > Result then
			Result := Excel_links[i];
	end;
end;

procedure LoadPreview;
var
	rows, cols, K:	integer;
	val1:			Extended;
begin
	SetLength(PreviewArray, 0);
	K := 0;
	
	for var index: integer := 0 to Form1.tbExcelTabs.Tabs.Count-1 do
	begin
		rows := VarArrayHighBound(ExcelArray[index], 1);
		cols := VarArrayHighBound(ExcelArray[index], 2);
		SetLength(PreviewArray, Length(PreviewArray) + rows);
	
		if 	(Excel_links[0] > -1) and
			(Excel_links[1] > -1) and
			(Excel_links[2] > -1) and
			(Excel_links[3] > -1) and
			(cols > GetMaxLink) then
		begin			
			for var i: integer := VarArrayLowBound(ExcelArray[index], 1) to rows do
			begin
				if  (String(ExcelArray[index][i, Excel_links[0]+1]) <> '') and
					(String(ExcelArray[index][i, Excel_links[1]+1]) <> '') and
					(TryStrToFloat(String(ExcelArray[index][i, Excel_links[2]+1]), val1)) and
					(TryStrToFloat(String(ExcelArray[index][i, Excel_links[3]+1]), val1)) then
				begin
					PreviewArray[K, 0] := String(ExcelArray[index][i, Excel_links[0]+1]);
					PreviewArray[K, 1] := String(ExcelArray[index][i, Excel_links[1]+1]);
					PreviewArray[K, 2] := String(ExcelArray[index][i, Excel_links[2]+1]);
					PreviewArray[K, 3] := String(ExcelArray[index][i, Excel_links[3]+1]);

					if cols > 4 then
					begin
						if ((Excel_links[4] > -1) and (String(ExcelArray[index][i, Excel_links[4]+1]) <> '')) then
							PreviewArray[K, 4] := String(ExcelArray[index][i, Excel_links[4]+1])
						else
							PreviewArray[K, 4] := '1';

						if (Excel_links[5] > -1) and (cols > 5) then
						begin
							PreviewArray[K, 5] := ParseYear(String(ExcelArray[index][i, Excel_links[5]+1]));
							if PreviewArray[K, 5] <> '' then
							begin
								PreviewArray[K, 0] := PreviewArray[K, 0] + ', ' + PreviewArray[K, 5] + '�.';
								PreviewArray[K, 1] := PreviewArray[K, 1] + '-' + PreviewArray[K, 5];
							end;
						end;

						if (Excel_links[6] > -1) and (cols > 6) then
						begin
							PreviewArray[K, 6] := String(ExcelArray[index][i, Excel_links[6]+1]);
							if PreviewArray[K, 6] <> '' then
								PreviewArray[K, 0] := PreviewArray[K, 0] + ', ' + PreviewArray[K, 6];
						end;

						if (Excel_links[7] > -1) and (cols > 7) then
						begin
							PreviewArray[K, 7] := String(ExcelArray[index][i, Excel_links[7]+1]);
							if PreviewArray[K, 7] <> '' then
								PreviewArray[K, 0] := PreviewArray[K, 0] + ', ' + PreviewArray[K, 7];
						end;
					end;
					Inc(K);
				end;
			end;
			SetLength(PreviewArray, K);
		end;
	end;
end;

procedure DrawPreview;
var
	N:	PVirtualNode;
begin
	Form1.PreviewTree.Clear;
    if Length(PreviewArray) > 0 then
    begin
        try
            Form1.PreviewTree.BeginUpdate;

            Form1.PreviewTree.RootNodeCount := Length(PreviewArray);

            N := Form1.PreviewTree.GetFirst;
            while Assigned(N) do
            begin
                N.SetData(pointer(N.Index));
                N := N.NextSibling;
        end
        finally
            Form1.PreviewTree.Header.AutoFitColumns(false);
            Form1.PreviewTree.EndUpdate;
            Form1.PreviewTree.Invalidate;
        end;
    end;
end;

procedure UploadPrice;
var
  So: 				tStringlist;
  regid:			AnsiString;
  buttonSelected:	Integer;
begin
	try
    	buttonSelected := MessageDlg('�� ������������� ������ ��������� ' + Length(PreviewArray).ToString +
        ' ������� ��� ������������� ' + Form1.cbDivisions.Text + '?', mtConfirmation, mbOKCancel, 0);

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
            So.SaveToFile('C:\ExcelParser\ExcelParser.sql');
            fcon.Execute(So.Text);

            ShowMessage('����� ������� ��������.');
        end;
    finally
    	So.Free;
    end;
end;
end.
