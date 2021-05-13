unit uxUtils;

interface

uses
  Types,
  Classes;

type
  tStrArr = array of AnsiString;

function IntRect(Rc:tRect;Ix:integer=1):tRect;overload;
function IntRect(Rc:tRect;L,T,R,B:integer):tRect;overload;
function SplitStr(Sx: AnsiString; var A: tStrArr): boolean;
function GetAlign(A: byte): TAlignment; overload;
function GetAlign(A: byte; SingleLine: boolean): WORD; overload;
function DateFromInt(V: string): AnsiString; overload;
function DateFromInt(V: integer): AnsiString; overload;
function iif(F: Boolean; A, B: string): string; overload;
function iif(F: Boolean; A, B: int64): int64; overload;
function Split(var S: AnsiString): AnsiString; overload;
function Split(S: AnsiString; var R: AnsiString): AnsiString; overload;
function SplitEx(var S: AnsiString; M: AnsiString = '.'): AnsiString; overload;
function IndexOf(S, Ñ: AnsiString): integer;
procedure SaveDebugInfo(S, Fn: string);
function IntToDate(Dx: integer): TDateTime;

const
  ctBlank = (-1);

function CmpCnt: integer;

var
  StopEvent: THandle;

  fLoopResult: boolean;
  StartMode : byte;


implementation

uses
  Windows,
  AnsiStrings,
  SysUtils;

var fCmpCnt: integer = 0;

function IntToDate(Dx: integer): TDateTime;
var D, M, Y: word;
begin
  D := Dx mod 100;
  Dx := Dx div 100;
  M := Dx mod 100;
  Y := Dx div 100;
  Result := EncodeDate(Y, M, D);
end;

function CmpCnt: integer;
begin
  Result := fCmpCnt;
  Inc(fCmpCnt);
end;

function DateFromInt(V: string): AnsiString; overload;
begin
  Result := Copy(V, 7, 2) + '.' + Copy(V, 5, 2) + '.' + Copy(V, 1, 4);
end;

function DateFromInt(V: integer): AnsiString; overload;
begin
  Result := DateFromInt(IntToStr(V));
end;

function IntRect(Rc: tRect; Ix: integer): tRect;
begin
  Result:=Rc;
  Inc(Result.Left,Ix);
  Inc(Result.Top,Ix);
  Dec(Result.Right,Ix);
  Dec(Result.Bottom,Ix);
end;

function IntRect(Rc: tRect; L, T, R, B: integer): tRect;
begin
  Result:=Rc;
  Inc(Result.Left,L);
  Inc(Result.Top,T);
  Dec(Result.Right,R);
  Dec(Result.Bottom,B);
end;

function SplitStr(Sx: AnsiString; var A: tStrArr): boolean;
var C: AnsiChar; S: AnsiString;
begin
  SetLength(A, 0);
  for C in Sx do
  begin
    case C of
      #32: if (Length(S)>0) then begin SetLength(A, Length(A) + 1); A[High(A)] := S; S := ''; end;
    else
      S := S + C;
    end;
  end;
  if (Length(S)>0) then
  begin
    SetLength(A, Length(A) + 1);
    A[High(A)] := S;
  end;
  Result := True;
end;

function GetAlign(A: byte): TAlignment;
const Al: array[1..9] of byte = (0, 2, 1, 0, 2, 1, 0, 2, 1);
begin
  Result := tAlignment(Al[A]);
end;

function GetAlign(A: byte; SingleLine: boolean): WORD;
var Dv, Dh: word;
begin
  Dv := 0; Dh:= 0;
  if A = 0 then
    Result := DT_LEFT + DT_WORDBREAK
  else
  begin
    case ((A-1) mod 3)+1 of
      1: Dh := DT_LEFT;
      2: Dh := DT_CENTER;
      3: Dh := DT_RIGHT;
    end;
    case ((A-1) div 3)+1 of
      1: Dv := DT_BOTTOM;
      2: Dv := DT_VCENTER;
      3: Dv := DT_TOP;
    end;
    Result := DT_SINGLELINE + DT_END_ELLIPSIS + Dh + Dv;
  end;
end;

function GetAlignEx(A: byte; SingleLine: boolean = True): WORD;
var Dv, Dh: word;
begin
  Dv := 0; Dh:= 0;
  if A = 0 then
    Result := DT_LEFT + DT_WORDBREAK
  else
  begin
    case ((A-1) mod 3)+1 of
      1: Dh := DT_LEFT;
      2: Dh := DT_CENTER;
      3: Dh := DT_RIGHT;
    end;
    case ((A-1) div 3)+1 of
      1: Dv := DT_BOTTOM;
      2: Dv := DT_VCENTER;
      3: Dv := DT_TOP;
    end;
    Result := DT_SINGLELINE + DT_END_ELLIPSIS + Dh + Dv;
  end;
end;

function iif(F: Boolean; A, B: string): string;
begin
  if F then Result := A else Result := B;
end;

function iif(F: Boolean; A, B: int64): int64;
begin
  if F then Result := A else Result := B;
end;

function Split(var S: AnsiString): AnsiString; overload;
var K: integer;
begin
  K := Pos('.', S);
  if (K>0) then
  begin
    Result := Copy(S, 1, K - 1);
    System.Delete(S, 1, K);
  end
  else
  begin
    Result := S;
    S := '';
  end;
end;

function SplitEx(var S: AnsiString; M: AnsiString = '.'): AnsiString; overload;
var K: integer;
begin
  K := Pos(M, S);
  if (K>0) then
  begin
    Result := Copy(S, 1, K - 1);
    System.Delete(S, 1, K);
  end
  else
  begin
    Result := S;
    S := '';
  end;
end;

function Split(S: AnsiString; var R: AnsiString): AnsiString; overload;
begin
  R := S;
  Result := Split(R);
end;

function IndexOf(S, Ñ: AnsiString): integer;
var Ps, Pc: pAnsiChar;
begin
  Result := 0;
  Pc := @Ñ[1];
  Ps := @S[1];
  while (Pc<>'') and (AnsiStrings.StrIComp(Pc, Ps)<>0) do
  begin
    Inc(Result);
    Inc(Pc, AnsiStrings.StrLen(Pc)+1);
  end;
  if (AnsiStrings.StrIComp(Pc, Ps)<>0) then Result := ctBlank;
end;

procedure SaveDebugInfo(S, Fn: string);
begin
  with TStringList.Create do
  try
    Text := S;
    SaveToFile(Fn);
  finally
    Free;
  end;
end;

end.
