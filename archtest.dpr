program archtest;

uses
  Forms,
  FmMain in 'FmMain.pas' {frmType},
  UFileCatcher in 'UFileCatcher.pas';

{$R *.res}
{$SetPEFlags 1}

begin
  Application.Initialize;
  Application.Title := 'Archive Test';
  Application.CreateForm(TfrmType, frmType);
  Application.Run;
end.
