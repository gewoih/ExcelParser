unit uxSQL;

interface

function ConnectSQL(Var Con: OleVariant; Suff: string = ''): boolean;
function GetConStr: String;

var
	fcon:	OleVariant;

implementation

uses System.Classes, System.Win.ComObj, System.SysUtils, uxParams;

function GetConStr: string;
begin
  With TStringList.Create do
  try
    LineBreak := ';';
    Values['Provider'] := fParams.AsStr['sql\Prov'];
    Values['Persist Security Info'] := 'False';
    Values['Data Source'] := fParams.AsStr['sql\Serv'];
    Values['Initial Catalog'] := fParams.AsStr['sql\Base'];
    Values['User ID'] := fParams.AsStr['sql\User'];
    Values['Application Name'] := ChangeFileExt(ExtractFileName(ParamStr(0)), '');
    Values['MultipleActiveResultSets'] := 'True';
    Values['Password'] := fParams.AsStr['sql\Pass'];
    Result := Text;
  finally
    Free;
  end;
end;

function ConnectSQL(Var Con: OleVariant; Suff: string = ''): boolean;
var
	S: AnsiString;
begin
  try
    Con := CreateOleObject('ADODB.Connection');
    Con.CursorLocation:= 3;
    Con.CommandTimeout := 60000;
    Con.ConnectionTimeout := 10;
    Con.Open(GetConStr);

    Result := True;
  except
    on E:Exception do
    begin
      S := E.Message;
      Result := False;
    end;
  end;
end;

end.
