program DummyImportOTM;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, datetimectrls, tachartlazaruspkg, uMain, uConvert, uSetProFinish, CreatePro;

{$R *.res}

begin
  RequireDerivedFormResource:=True;
  Application.Scaled:=True;
  Application.Initialize;
  Application.CreateForm(TfMain, fMain);
  Application.CreateForm(TfConvert, fConvert);
  Application.CreateForm(TFormSetProFinish, FormSetProFinish);
  Application.Run;
end.

