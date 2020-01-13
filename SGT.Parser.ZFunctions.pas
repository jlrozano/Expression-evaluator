unit SGT.Parser.ZFunctions;

interface

uses SGT.Parser.ZParser, SysUtils, System.Variants;

implementation

// Get current date and time
function Now_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 0, 'Now');
  Result := Now
end;

function FormatNum_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 2, 'FormatNum');
  if Result<>2 then
    EParseException.CreateFmt(rsInvalidParamsNo,['FormatNum']);
  Result := FormatFloat(Args[0], Args[1]);
end;

function Fcur_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'FCur');
  Result := FormatFloat(',##0.00 €',Args[0]);
end;

function IsNull_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'IsNull');
  Result := VarIsNull(Args[0]) Or VarIsEmpty(Args[0]);
end;

function FNum_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'FNum');
  Result := Format('%n',[Args[0]]);
end;

function StrPos_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 2, 'StrPos');
  Result := Pos(Args[0],Args[1]);
end;

function Replace_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 3, 'Replace');
  Result := StringReplace(Args[0],Args[1],Args[2],[rfIgnoreCase])
end;

function ReplaceAll_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 3, 'ReplaceAll');
  Result := StringReplace(Args[0],Args[1],Args[2],[rfIgnoreCase, rfReplaceAll])
end;


function StrCopy_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 3, 'StrCopy');
  Result := Copy(Args[0], Integer(Args[1]), Integer(Args[2]));
end;

function StrLen_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'StrLen');
  Result := Length(Args[0]);
end;

function IIf_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 3, 'Iif');
  if Boolean(VarAsType(Args[0], varBoolean)) then result:= Args[1]
  else result:= Args[2];
end;



function Upper_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'Upper');
  result:=AnsiUpperCase(Args[0]);
end;

function Lower_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'Lower');
  result:=AnsiLowerCase(Args[0]);
end;



function Date_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'Date');
  Result := Date //FormatDateTime('yyyy-mm-dd hh:nn:ss', Now());
end;

function FormatDate_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 2, 'FomatDate');
  result:= FormatDateTime(Args[0], Args[1])
end;

{ function GetItem_Execute(Sender: TParser; const Args: Array of Variant): Variant;
var
  LList: IListInterface;
  LValue: TValue;
begin
  CheckParamCount(Args, 2, 'GetListItem');
  if VarType(Args[0])= varInteger then
    LValue:= TObject(Integer(Args[0]))
  else
    LValue:= TValue.FromVariant(Args[0]);
  LList:= TListInterface.Get(Lvalue);
  if LList=nil then
    raise Exception.Create('Error GetItem function. Parameter not is a list');

  try
    LValue:= LList.GetItem(Args[1]);
    result:= LValue.AsVariant;
  finally
    LList.DetachList;
  end;
end;

function GetListCount_Execute(Sender: TParser; const Args: Array of Variant): Variant;
var
  LList: IListInterface;
  LValue: TValue;
begin
  CheckParamCount(Args, 1, 'GetListSize');
  if VarType(Args[0])= varInteger then
    LValue:= TObject(Integer(Args[0]))
  else
    LValue:= TValue.FromVariant(Args[0]);
  LList:= TListInterface.Get(Lvalue);
  if LList=nil then
    raise Exception.Create('Error GetListSize function. Parameter not is a list');
  try
    result:= LList.Size;
  finally
    LList.DetachList;
  end;
end;
 }

initialization
  TParser.AddFunction('now',Now_Execute);
  TParser.AddFunction('iif',Iif_Execute);
  TParser.AddFunction('strcopy',StrCopy_Execute);
  TParser.AddFunction('upper',Upper_Execute);
  TParser.AddFunction('lower',lower_Execute);
  TParser.AddFunction('date',Date_Execute);
  TParser.AddFunction('formatdate',FormatDate_Execute);
  TParser.AddFunction('replace',Replace_Execute);
  TParser.AddFunction('RepalceAll',ReplaceAll_Execute);
  TParser.AddFunction('strpos', StrPos_Execute);
  TParser.AddFunction('FormatNum',FormatNum_Execute);
  TParser.AddFunction('fcur',FCur_Execute);
  TParser.AddFunction('fnum',FNum_Execute);
  TParser.AddFunction('isnull',IsNull_Execute);
  TParser.AddFunction('strlen',StrLen_Execute);
//  TParser.AddFunction('getlistitem',GetItem_Execute);
//  TParser.AddFunction('getlistsize',GetListCount_Execute);

end.
