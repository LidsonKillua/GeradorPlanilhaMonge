program PlanilhaMonge;

uses
  Vcl.Forms,
  PMPrincipal in 'PMPrincipal.pas' {frmGerarPlanilhaPayback};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmGerarPlanilhaPayback, frmGerarPlanilhaPayback);
  Application.Run;
end.
