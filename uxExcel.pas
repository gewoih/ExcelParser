unit uxExcel;

interface

uses VirtualTrees;

procedure AddExcel;
procedure DrawExcel;

var
	ExcelArray:	array of OleVariant;

implementation

uses MainForm, System.Variants, System.Classes, SysUtils, uxSuppliers, Dialogs, System.Win.ComObj,
uxExcelLinks, uxPreview, uxDivisions;

procedure DrawExcel;
procedure SetTreeColumns;
var
	Column:		TVirtualTreeColumn;
begin
	try
    	Form1.ExcelTree.BeginUpdate;
        Form1.ExcelTree.Header.Columns.Clear;
        
        for var i := VarArrayLowBound(ExcelArray[tab_index], 2)-1 to VarArrayHighBound(ExcelArray[tab_index], 2)-1 do
        begin
            Column := Form1.ExcelTree.Header.Columns.Add;
            Column.Alignment := taLeftJustify;
            Column.CaptionAlignment := taCenter;
            Column.Text := i.ToString;
        end;
    finally
    	Form1.ExcelTree.EndUpdate;
        Form1.ExcelTree.Invalidate;	
    end;
end;

var
	N:	PVirtualNode;
begin
    try
    	Form1.ExcelTree.BeginUpdate;
    	Form1.ExcelTree.Clear;
        
        SetTreeColumns;

        Form1.ExcelTree.RootNodeCount := VarArrayHighBound(ExcelArray[tab_index], 1);
        
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

procedure AddExcel;
var
	Excel, Sheet:	OleVariant;
begin
    if Assigned(Form1.SuppliersTree.FocusedNode) and Form1.OpenDialog1.Execute then
    begin
        try
            Form1.tbExcelTabs.Tabs.Clear;
            
            Excel := CreateOleObject('Excel.Application');
            Excel.Workbooks.Open(Form1.OpenDialog1.FileName, 0, true);
 
            SetLength(ExcelArray, integer(Excel.ActiveWorkBook.Sheets.Count));
            for var i: integer := 1 to Excel.ActiveWorkBook.Sheets.Count do
            begin
            	Form1.tbExcelTabs.Tabs.Add(Excel.ActiveWorkBook.Sheets.Item[i].Name);
            	Sheet := Excel.ActiveWorkBook.Sheets.Item[i];

            	ExcelArray[i-1] := Sheet.UsedRange.Value;
            end;
			Form1.tbExcelTabs.TabIndex := 0;
			
			DrawLinks;
			DrawDivisions;
			
			LoadPreview;
			DrawPreview;
        finally
        	Sheet := Unassigned;
            Excel.Quit;
            Excel := 0;
        end;
    end;
end;

end.
