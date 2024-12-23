unit PMPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons;

type
  TfrmGerarPlanilhaPayback = class(TForm)
    Button1: TButton;
    OpenDialog: TOpenDialog;
    edtPlanilhaBase: TEdit;
    btnBrowse: TSpeedButton;
    edtPdcAnual: TEdit;
    edtConsAnual: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    edtInvestInic: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    edtCidade: TEdit;
    Label7: TLabel;
    edtArquivo: TEdit;
    Label8: TLabel;
    edtModulos: TEdit;
    Label9: TLabel;
    edtKWp: TEdit;
    Label10: TLabel;
    edtArea: TEdit;
    procedure Button1Click(Sender: TObject);
    procedure btnBrowseClick(Sender: TObject);
  private
    procedure ValidarCampos;
  public
    { Public declarations }
  end;

var
  frmGerarPlanilhaPayback: TfrmGerarPlanilhaPayback;

implementation

{$R *.dfm}

uses
  ComObj;

procedure TfrmGerarPlanilhaPayback.btnBrowseClick(Sender: TObject);
begin
  if OpenDialog.Execute then
    edtPlanilhaBase.Text := OpenDialog.FileName;
end;

procedure TfrmGerarPlanilhaPayback.Button1Click(Sender: TObject);
var
  Excel: Variant;
  Linha: Integer;
begin
  Self.ValidarCampos();

  Excel := CreateOleObject('Excel.Application');
  Excel.Visible := True;
  Excel.WorkBooks.Open(edtPlanilhaBase.Text);
  Excel.WorkBooks[1].Sheets[1].Cells[2, 9]:= edtInvestInic.Text;
  Excel.WorkBooks[1].Sheets[1].Cells[2, 1]:= edtCidade.Text;
  Excel.WorkBooks[1].Sheets[1].Cells[3, 1]:= edtArquivo.Text;
  Excel.WorkBooks[1].Sheets[1].Cells[4, 1]:= edtModulos.Text + ' módulos';
  Excel.WorkBooks[1].Sheets[1].Cells[5, 1]:= edtKWp.Text + ' KWp';
  Excel.WorkBooks[1].Sheets[1].Cells[6, 1]:= edtArea.Text + ' m²';

  for Linha := 3 to 27 do
  begin
    Excel.WorkBooks[1].Sheets[1].Cells[Linha, 3] := edtPdcAnual.Text;
    Excel.WorkBooks[1].Sheets[1].Cells[Linha, 4] := edtConsAnual.Text;
  end;

  Excel.WorkBooks[1].SaveAs(ExtractFilePath(edtPlanilhaBase.Text) + edtCidade.Text + '.xlsx');
end;

procedure TfrmGerarPlanilhaPayback.ValidarCampos;

  procedure VerifEditVazio(edt: TEdit);
  begin
    if Trim(edt.Text) = '' then
    begin
      if edt.CanFocus then
        edt.SetFocus;

      raise Exception.Create('Campo não preenchido: ' + edt.Hint);
    end;
  end;

begin
  VerifEditVazio(edtPlanilhaBase);
  VerifEditVazio(edtPdcAnual);
  VerifEditVazio(edtConsAnual);
  VerifEditVazio(edtInvestInic);
  VerifEditVazio(edtCidade);
  VerifEditVazio(edtArquivo);
  VerifEditVazio(edtModulos);
  VerifEditVazio(edtKWp);
  VerifEditVazio(edtArea);
end;

end.
