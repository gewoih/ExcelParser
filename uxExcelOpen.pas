unit uxExcelOpen;

interface

implementation

procedure TForm1.DrawPreview(Rows, Columns: integer);
procedure SetTreeColumns;
var
	Column:	TVirtualTreeColumn;
begin
	try
        for var i: integer := 0 to Columns-1 do
        begin
            Column := PreviewTree.Header.Columns.Add;
            Column.Alignment := taLeftJustify;
            Column.Text := (i+1).ToString;
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
        SetTreeColumns;

        PreviewTree.RootNodeCount := Rows;

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

end.
