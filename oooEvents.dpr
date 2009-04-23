program oooEvents;

uses
  Forms,
  oooEvents1 in 'oooEvents1.pas' {Form1},
  oooEvents_TLB in 'oooEvents_TLB.pas',
  oooEvents2 in 'oooEvents2.pas' {Listener: CoClass};

{$R *.TLB}

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
