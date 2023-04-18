program PythonComTest;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  System.SysUtils, System.Win.ComObj, Winapi.ActiveX;

var
  ComObj: OleVariant;
  PersonObj: OleVariant;
  PersonStr: string;

begin
  CoInitialize(nil);

  try
    try
      ComObj := CreateOleObject('Python.MyCOMObject');
      PersonObj := ComObj.get_person();
      PersonStr := Format('%s (%d years old)', [PersonObj.name, PersonObj.age]);
      WriteLn(PersonStr);

      WriteLn('Press Enter to exit...');
      ReadLn;
    except
      on E: Exception do
        WriteLn(E.ClassName, ': ', E.Message);
    end;
  finally
    CoUninitialize;
  end;

end.
