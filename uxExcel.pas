unit uxExcel;

interface

uses VirtualTrees;

procedure AddExcel;
//procedure LoadExcel;
procedure DrawExcel;
//procedure UploadExcel;

var
	ExcelArray:	OleVariant;

implementation

uses MainForm, System.Variants, System.Classes, SysUtils, uxSuppliers, Dialogs, System.Win.ComObj,
uxExcelLinks, uxPreview;

procedure DrawExcel;
procedure SetTreeColumns;
var
	Column:	TVirtualTreeColumn;
begin
	try
        Form1.ExcelTree.Header.Columns.Clear;

        for var i: integer := 0 to VarArrayHighBound(ExcelArray, 2)-1 do
        begin
            Column := Form1.ExcelTree.Header.Columns.Add;
            Column.Alignment := taLeftJustify;
            Column.CaptionAlignment := taCenter;
            Column.Text := i.ToString;
        end;
    finally
        Form1.ExcelTree.Invalidate;
    end;
end;

var
	N:		PVirtualNode;
    row:	integer;
begin
    try
        Form1.ExcelTree.BeginUpdate;
        Form1.ExcelTree.Clear;

        SetTreeColumns;

        Form1.ExcelTree.RootNodeCount := VarArrayHighBound(ExcelArray, 1);

        N := Form1.ExcelTree.GetFirst;
        while Assigned(N) do
        begin
            N.SetData(pointer(N.Index));
            N := N.NextSibling;
        end
    finally
        Form1.ExcelTree.Header.AutoFitColumns(false);
        Form1.ExcelTree.EndUpdate;
        Form1.ExcelTree.Invalidate;
    end;
end;

//procedure LoadExcel;
//var
//	R:		   			OleVariant;
//    rows, cols, index:	integer;
//begin
//	index := Suppliers[Form1.SuppliersTree.FocusedNode.Index].id;
//
//    R := fcon.Execute(Format(Form1.scLoadExcel.Items.Text, [index]));
//
//    rows := R.RecordCount;
//    cols := R.Fields.Count;
//
//    ExcelArray := VarArrayCreate([0, rows, 0, cols], varVariant);
//
//    while not R.EOF do
//    begin
//        for var i: integer := 0 to cols-1 do
//        begin
//            ExcelArray[R.AbsolutePosition, i] := R.Fields[i].Value;
//        end;
//        R.MoveNext;
//    end;
//end;

//procedure UploadExcel;
//var
//	id, Rows, Columns:	integer;
//    List:				TStringList;
//    S:					String;
//begin
//    try
//    	if VarArrayHighBound(ExcelArray, 1) >= 50 then
//        begin
//            id := Suppliers[Form1.SuppliersTree.FocusedNode.Index].id;
//            List := TStringList.Create;
//
//            Rows := 50;
//            Columns := VarArrayHighBound(ExcelArray, 2);
//
//            fcon.Execute('delete from Excel_templates where row > 0 and linkid = ' + id.ToString);
//
//            List.Add('insert into Excel_templates values');
//            for var i: integer := 1 to Rows do
//            begin
//                for var j: integer := 1 to Columns do
//                begin
//                    //if not VarIsClear(ExcelArray[i, j]) then
//                    begin
//                        if VarType(ExcelArray[i, j]) = varDouble then
//                            S := QuotedStr(extended(ExcelArray[i, j]).ToString)
//                        else
//                            S := QuotedStr(ExcelArray[i, j]);
//
//                        List.Add(Format('(' + id.ToString + ',%d, %d, %s),', [i, j, S]));
//                    end;
//                end;
//            end;
//
//            List.Add('(' + id.ToString + ',0, 1, null),');
//            List.Add('(' + id.ToString + ',0, 2, null),');
//            List.Add('(' + id.ToString + ',0, 3, null),');
//            List.Add('(' + id.ToString + ',0, 4, null)');
//
//            fcon.Execute(List.Text);
//
//            LoadExcel;
//            DrawExcel;
//        end
//        else
//        	ShowMessage('Выберите файл в котором больше 50 строк!');
//    finally
//
//    end;
//end;

procedure AddExcel;
var
	Excel, Sheet:	OleVariant;
begin
    if Form1.OpenDialog1.Execute and Assigned(Form1.SuppliersTree.FocusedNode) then
    begin
        try
        	SetLength(PreviewArray, 0);

            Excel := CreateOleObject('Excel.Application');
            Excel.Workbooks.Open(Form1.OpenDialog1.FileName, 0, true);
            Sheet := Excel.ActiveWorkbook.ActiveSheet;

            ExcelArray := Sheet.UsedRange.Value;

            DrawExcel;
            DrawLinks;

            LoadPreview;
            DrawPreview;
        finally
        	if VarIsEmpty(Excel) = false then
  			begin
    			Excel.Quit;
    			Excel := 0;
			end;
        end;
    end;
end;

end.
