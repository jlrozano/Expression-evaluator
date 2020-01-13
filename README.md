# Expression-evaluator
A Delphi simple expression evaluator based with user functions and variables. This work is based in old Zeos parse libaries. 

Support varaibles and funcions definions.

# Operators

Supported operators:

* + 
* -
* =
* >=
* <=
* % (mod)
* IN (value in array values. Array is supported with [])
* AND
* OR
* XOR
* NOT 
* LIKE (see SGT.Parse.Match for details)

# Simple usage

Begin
  Parser = TParser.Create();
  P.Equation = '(5*7) / 2)'
  WriteLn(P.Evalute<Int>());
End

# Usage with variables

procedure GetVar(sender: TParser; VarName: string; var Value: Variant);
begin
   If (UpperCase(VarName) = 'X') Then
     Value := 3
   else
     Raise Exception.Create('Unkonow variable');
end;

Begin
  Parser = TParser.Create(GetVar);
  P.Equation = '(5*7) / x)'
  WriteLn(P.Evalute<Int>());
End

# Define custom functions

See SGT.Parser.Function for examples.

procedure GetVar(sender: TParser; VarName: string; var Value: Variant);
begin
   // In a form with a dataset, you can use Value:= DataSet.FieldByName(VarName).Value
   If (UpperCase(VarName) = 'X') Then
     Value := 'say hello!!'
   else
     Raise Exception.Create('Unkonow variable');
end;

begin
    p:= TParser.Create(GetVar) ;
    TParser.AddFunction('Upper', function (sender: TParser; const Args: Array of Variant): Variant
    begin
      result := UpperCase(Args[0]);
    end
    );
    p.Equation := 'Upper(x)';
    Writeln(p.Evalute<String>);
 end;  
