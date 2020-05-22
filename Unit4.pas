unit Unit4;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ExtCtrls, StdCtrls;

type
  TForm4 = class(TForm)
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;

implementation

uses ProgramAnalys, Unit2, Unit3;

{$R *.dfm}

procedure TForm4.N1Click(Sender: TObject);
begin
  Form4.Close;
  If (Form1.Visible=False) and (Form2.Visible=True) then Form2.Show()
  else
  If (Form1.Visible=False) and (Form3.Visible=True) then Form3.Show()
  else
  Form1.Show();
end;

procedure TForm4.N2Click(Sender: TObject);
begin
 Form4.Close;
 Form1.close;
end;

end.
