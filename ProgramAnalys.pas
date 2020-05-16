unit ProgramAnalys;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, ComObj;

type
  TForm1 = class(TForm)
    StringGrid1: TStringGrid;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    OpenDialog1: TOpenDialog;
    Button4: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Excel: Variant;

implementation

uses Unit2, Unit3;
procedure Import_Data(XLSFile:string; Grid:TStringGrid);
 const
  xlCellTypeLastCell=$0000000B;
Var
 ExlApp, Sheet: OLEVariant;
 i,j,r,c:integer;

begin
ExlApp:=CreateOleObject('Excel.Application');
ExlApp.Visible:=false;
ExlApp.Workbooks.Open(XLSFile);
Sheet:=ExlApp.Workbooks[ExtractFileName(XLSFile)].WorkSheets[1];
Sheet.Cells.SpecialCells(xlCellTypeLastCell,EmptyParam).Activate;
r:=ExlApp.ActiveCell.Row;
c:=ExlApp.ActiveCell.Column;
Grid.RowCount:=r;
Grid.ColCount:=c;
for j:=1 to r do
 for i:=1 to c do
  Grid.Cells[i-1,j-1]:=sheet.cells[j,i];

ExlApp.Quit;
ExlApp:=Unassigned;
Sheet:=Unassigned;
end;

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
If OpenDialog1.Execute then Import_Data(OpenDialog1.FileName,StringGrid1);
end;

procedure TForm1.Button3Click(Sender: TObject);
Var i,j:integer;
    k,l,m:double;
begin
  Form2.StringGrid.ColCount:=6;
  Form2.StringGrid.RowCount:=6;
  Form2.StringGrid.ColWidths[0] := 280;
  for i:=1 to Form2.StringGrid.ColCount-1 do
   for j:=1 to Form2.StringGrid.RowCount-1 do
    Form2.StringGrid.Cells[0,0]:='Показатель';
    Form2.StringGrid.Cells[0,1]:='Коэффициент Абсолютной Ликвидности';
    Form2.StringGrid.Cells[0,2]:='Коэффициент Промежуточной Ликвидности';
    Form2.StringGrid.Cells[0,3]:='Коэффициент Текущей Ликвидности';
    Form2.StringGrid.Cells[0,4]:='Коэффициент Обеспеченности СОС';
    Form2.StringGrid.Cells[0,5]:='Коэффициент Восстановления Платежеспособности';
    Form2.StringGrid.Cells[1,0]:='2016';
    Form2.StringGrid.Cells[2,0]:='2017';
    Form2.StringGrid.Cells[3,0]:='2018';
    Form2.StringGrid.Cells[4,0]:='2017-2016';
    Form2.StringGrid.Cells[5,0]:='2018-2017';
    k:=((StrToInt(StringGrid1.Cells[1,2])+StrToInt(Form1.StringGrid1.Cells[1,3])) / (StrToInt(Form1.StringGrid1.Cells[1,4])+StrToInt(Form1.StringGrid1.Cells[1,5])));
    Form2.StringGrid.Cells[1,1]:=FloatToStrF(k,fffixed,19,2);
    l:=(StrToInt(Form1.StringGrid1.Cells[2,2])+StrToInt(Form1.StringGrid1.Cells[2,3])) / (StrToInt(Form1.StringGrid1.Cells[2,4])+StrToInt(Form1.StringGrid1.Cells[2,5]));
    Form2.StringGrid.Cells[2,1]:=FloatToStrF(l,fffixed,19,2);
    m:=(StrToInt(Form1.StringGrid1.Cells[3,2])+StrToInt(Form1.StringGrid1.Cells[3,3])) / (StrToInt(Form1.StringGrid1.Cells[3,4])+StrToInt(Form1.StringGrid1.Cells[3,5]));
    Form2.StringGrid.Cells[3,1]:=FloatToStrF(m,fffixed,19,2);
    Form2.StringGrid.Cells[4,1]:=FloatToStrF(l-k,fffixed,19,2);
    Form2.StringGrid.Cells[5,1]:=FloatToStrF(m-l,fffixed,19,2);
    k:=((StrToInt(StringGrid1.Cells[1,6])+StrToInt(Form1.StringGrid1.Cells[1,2])+StrToInt(Form1.StringGrid1.Cells[1,3])) / (StrToInt(Form1.StringGrid1.Cells[1,4])+StrToInt(Form1.StringGrid1.Cells[1,5])));
    Form2.StringGrid.Cells[1,2]:=FloatToStrF(k,fffixed,19,2);
    l:=((StrToInt(StringGrid1.Cells[2,6])+StrToInt(Form1.StringGrid1.Cells[2,2])+StrToInt(Form1.StringGrid1.Cells[2,3])) / (StrToInt(Form1.StringGrid1.Cells[2,4])+StrToInt(Form1.StringGrid1.Cells[2,5])));
    Form2.StringGrid.Cells[2,2]:=FloatToStrF(l,fffixed,19,2);
    m:=((StrToInt(StringGrid1.Cells[3,6])+StrToInt(Form1.StringGrid1.Cells[3,2])+StrToInt(Form1.StringGrid1.Cells[3,3])) / (StrToInt(Form1.StringGrid1.Cells[3,4])+StrToInt(Form1.StringGrid1.Cells[3,5])));
    Form2.StringGrid.Cells[3,2]:=FloatToStrF(m,fffixed,19,2);
    Form2.StringGrid.Cells[4,2]:=FloatToStrF(l-k,fffixed,19,2);
    Form2.StringGrid.Cells[5,2]:=FloatToStrF(m-l,fffixed,19,2);
    k:=StrToInt(Form1.StringGrid1.Cells[1,7]) / (StrToInt(Form1.StringGrid1.Cells[1,4])+StrToInt(Form1.StringGrid1.Cells[1,5]));
    Form2.StringGrid.Cells[1,3]:=FloatToStrF(k,fffixed,19,2);
    l:=StrToInt(Form1.StringGrid1.Cells[2,7]) / (StrToInt(Form1.StringGrid1.Cells[2,4])+StrToInt(Form1.StringGrid1.Cells[2,5]));
    Form2.StringGrid.Cells[2,3]:=FloatToStrF(l,fffixed,19,2);
    m:=StrToInt(Form1.StringGrid1.Cells[3,7]) / (StrToInt(Form1.StringGrid1.Cells[3,4])+StrToInt(Form1.StringGrid1.Cells[3,5]));
    Form2.StringGrid.Cells[3,3]:=FloatToStrF(m,fffixed,19,2);
    Form2.StringGrid.Cells[4,3]:=FloatToStrF(l-k,fffixed,19,2);
    Form2.StringGrid.Cells[5,3]:=FloatToStrF(m-l,fffixed,19,2);
    k:=(StrToInt(Form1.StringGrid1.Cells[1,13])+StrToInt(Form1.StringGrid1.Cells[1,14])-StrToInt(Form1.StringGrid1.Cells[1,11])) / StrToInt(Form1.StringGrid1.Cells[1,12]);
    Form2.StringGrid.Cells[1,4]:=FloatToStrF(k,fffixed,19,2);
    l:=(StrToInt(Form1.StringGrid1.Cells[2,13])+StrToInt(Form1.StringGrid1.Cells[2,14])-StrToInt(Form1.StringGrid1.Cells[2,11])) / StrToInt(Form1.StringGrid1.Cells[2,12]);
    Form2.StringGrid.Cells[2,4]:=FloatToStrF(l,fffixed,19,2);
    m:=(StrToInt(Form1.StringGrid1.Cells[3,13])+StrToInt(Form1.StringGrid1.Cells[3,14])-StrToInt(Form1.StringGrid1.Cells[3,11])) / StrToInt(Form1.StringGrid1.Cells[3,12]);
    Form2.StringGrid.Cells[3,4]:=FloatToStrF(m,fffixed,19,2);
    Form2.StringGrid.Cells[4,4]:=FloatToStrF(l-k,fffixed,19,2);
    Form2.StringGrid.Cells[5,4]:=FloatToStrF(m-l,fffixed,19,2);
    Form2.StringGrid.Cells[1,5]:='Не считается';
    l:=(StrToFloat(Form2.StringGrid.Cells[2,4])+(StrToFloat(Form2.StringGrid.Cells[3,4])-StrToFloat(Form2.StringGrid.Cells[2,4]))/4)/2;
    Form2.StringGrid.Cells[2,5]:=FloatToStrF(l,fffixed,19,2);
    m:=(StrToFloat(Form2.StringGrid.Cells[3,4])+(StrToFloat(Form2.StringGrid.Cells[4,4])-StrToFloat(Form2.StringGrid.Cells[3,4]))/4)/2;
    Form2.StringGrid.Cells[3,5]:=FloatToStrF(m,fffixed,19,2);
    Form2.StringGrid.Cells[4,5]:='Не считается';
    Form2.StringGrid.Cells[5,5]:=FloatToStrF(m-l,fffixed,19,2);
 Form2.StringGrid.ColWidths[1]:=75;
 Form2.StringGrid.ColWidths[4]:=75;
 Form2.Show();
 Form1.Visible:=false;
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
 Form1.Close;
end;

procedure TForm1.Button2Click(Sender: TObject);
Var i,j:integer;
    k,l,m:double;
begin
  Form3.StringGrid1.ColCount:=6;
  Form3.StringGrid1.RowCount:=8;
  Form3.StringGrid1.ColWidths[0] := 315;
  for i:=1 to Form3.StringGrid1.ColCount-1 do
   for j:=1 to Form3.StringGrid1.RowCount-1 do
    Form3.StringGrid1.Cells[0,0]:='Показатель';
    Form3.StringGrid1.Cells[0,1]:='Коэффициент Автономности';
    Form3.StringGrid1.Cells[0,2]:='Коэффициент Зависимости';
    Form3.StringGrid1.Cells[0,3]:='Коэффициент Финансовой Устойчивости';
    Form3.StringGrid1.Cells[0,4]:='Коэффициент Финансовой Активности';
    Form3.StringGrid1.Cells[0,5]:='Коэффициент Долгосрочного Привлечения';
    Form3.StringGrid1.Cells[0,6]:='Коэффициент Мобильности СОС';
    Form3.StringGrid1.Cells[0,7]:='Коэффициент Имущества Производственного Назначения';
    Form3.StringGrid1.Cells[1,0]:='2016';
    Form3.StringGrid1.Cells[2,0]:='2017';
    Form3.StringGrid1.Cells[3,0]:='2018';
    Form3.StringGrid1.Cells[4,0]:='2017-2016';
    Form3.StringGrid1.Cells[5,0]:='2018-2017';
    k:=StrToFloat(Form1.StringGrid1.Cells[1,13])/StrToFloat(Form1.StringGrid1.Cells[1,16]);
    Form3.StringGrid1.Cells[1,1]:=FloatToStrF(k,fffixed,19,4);
    l:=StrToFloat(Form1.StringGrid1.Cells[2,13])/StrToFloat(Form1.StringGrid1.Cells[2,16]);
    Form3.StringGrid1.Cells[2,1]:=FloatToStrF(l,fffixed,19,4);
    m:=StrToFloat(Form1.StringGrid1.Cells[3,13])/StrToFloat(Form1.StringGrid1.Cells[3,16]);
    Form3.StringGrid1.Cells[3,1]:=FloatToStrF(m,fffixed,19,4);
    Form3.StringGrid1.Cells[4,1]:=FloatToStrF(l-k,fffixed,19,4);
    Form3.StringGrid1.Cells[5,1]:=FloatToStrF(m-l,fffixed,19,4);
    k:=(StrToFloat(Form1.StringGrid1.Cells[1,14])+StrToFloat(Form1.StringGrid1.Cells[1,15]))/StrToFloat(Form1.StringGrid1.Cells[1,16]);
    Form3.StringGrid1.Cells[1,2]:=FloatToStrF(k,fffixed,19,4);
    l:=(StrToFloat(Form1.StringGrid1.Cells[2,14])+StrToFloat(Form1.StringGrid1.Cells[2,15]))/StrToFloat(Form1.StringGrid1.Cells[2,16]);
    Form3.StringGrid1.Cells[2,2]:=FloatToStrF(l,fffixed,19,4);
    m:=(StrToFloat(Form1.StringGrid1.Cells[3,14])+StrToFloat(Form1.StringGrid1.Cells[3,15]))/StrToFloat(Form1.StringGrid1.Cells[3,16]);
    Form3.StringGrid1.Cells[3,2]:=FloatToStrF(m,fffixed,19,4);
    Form3.StringGrid1.Cells[4,2]:=FloatToStrF(l-k,fffixed,19,4);
    Form3.StringGrid1.Cells[5,2]:=FloatToStrF(m-l,fffixed,19,4);
    k:=(StrToFloat(Form1.StringGrid1.Cells[1,13])+StrToFloat(Form1.StringGrid1.Cells[1,14]))/StrToFloat(Form1.StringGrid1.Cells[1,16]);
    Form3.StringGrid1.Cells[1,3]:=FloatToStrF(k,fffixed,19,4);
    l:=(StrToFloat(Form1.StringGrid1.Cells[2,13])+StrToFloat(Form1.StringGrid1.Cells[2,14]))/StrToFloat(Form1.StringGrid1.Cells[2,16]);
    Form3.StringGrid1.Cells[2,3]:=FloatToStrF(l,fffixed,19,4);
    m:=(StrToFloat(Form1.StringGrid1.Cells[3,13])+StrToFloat(Form1.StringGrid1.Cells[3,14]))/StrToFloat(Form1.StringGrid1.Cells[3,16]);
    Form3.StringGrid1.Cells[3,3]:=FloatToStrF(m,fffixed,19,4);
    Form3.StringGrid1.Cells[4,3]:=FloatToStrF(l-k,fffixed,19,4);
    Form3.StringGrid1.Cells[5,3]:=FloatToStrF(m-l,fffixed,19,4);
    k:=(StrToFloat(Form1.StringGrid1.Cells[1,14])+StrToFloat(Form1.StringGrid1.Cells[1,15]))/StrToFloat(Form1.StringGrid1.Cells[1,13]);
    Form3.StringGrid1.Cells[1,4]:=FloatToStrF(k,fffixed,19,4);
    l:=(StrToFloat(Form1.StringGrid1.Cells[2,14])+StrToFloat(Form1.StringGrid1.Cells[2,15]))/StrToFloat(Form1.StringGrid1.Cells[2,13]);
    Form3.StringGrid1.Cells[2,4]:=FloatToStrF(l,fffixed,19,4);
    m:=(StrToFloat(Form1.StringGrid1.Cells[3,14])+StrToFloat(Form1.StringGrid1.Cells[3,15]))/StrToFloat(Form1.StringGrid1.Cells[3,13]);
    Form3.StringGrid1.Cells[3,4]:=FloatToStrF(m,fffixed,19,4);
    Form3.StringGrid1.Cells[4,4]:=FloatToStrF(l-k,fffixed,19,4);
    Form3.StringGrid1.Cells[5,4]:=FloatToStrF(m-l,fffixed,19,4);
    k:=StrToFloat(Form1.StringGrid1.Cells[1,14])/(StrToFloat(Form1.StringGrid1.Cells[1,14])+StrToFloat(Form1.StringGrid1.Cells[1,15]));
    Form3.StringGrid1.Cells[1,5]:=FloatToStrF(k,fffixed,19,4);
    l:=StrToFloat(Form1.StringGrid1.Cells[2,14])/(StrToFloat(Form1.StringGrid1.Cells[2,14])+StrToFloat(Form1.StringGrid1.Cells[2,15]));
    Form3.StringGrid1.Cells[2,5]:=FloatToStrF(l,fffixed,19,4);
    m:=StrToFloat(Form1.StringGrid1.Cells[3,14])/(StrToFloat(Form1.StringGrid1.Cells[3,14])+StrToFloat(Form1.StringGrid1.Cells[3,15]));
    Form3.StringGrid1.Cells[3,5]:=FloatToStrF(m,fffixed,19,4);
    Form3.StringGrid1.Cells[4,5]:=FloatToStrF(l-k,fffixed,19,4);
    Form3.StringGrid1.Cells[5,5]:=FloatToStrF(m-l,fffixed,19,4);
    k:=(StrToFloat(Form1.StringGrid1.Cells[1,12])-StrToFloat(Form1.StringGrid1.Cells[1,15]))/StrToFloat(Form1.StringGrid1.Cells[1,13]);
    Form3.StringGrid1.Cells[1,6]:=FloatToStrF(k,fffixed,19,4);
    l:=(StrToFloat(Form1.StringGrid1.Cells[2,12])-StrToFloat(Form1.StringGrid1.Cells[2,15]))/StrToFloat(Form1.StringGrid1.Cells[2,13]);
    Form3.StringGrid1.Cells[2,6]:=FloatToStrF(l,fffixed,19,4);
    m:=(StrToFloat(Form1.StringGrid1.Cells[3,12])-StrToFloat(Form1.StringGrid1.Cells[3,15]))/StrToFloat(Form1.StringGrid1.Cells[3,13]);
    Form3.StringGrid1.Cells[3,6]:=FloatToStrF(m,fffixed,19,4);
    Form3.StringGrid1.Cells[4,6]:=FloatToStrF(l-k,fffixed,19,4);
    Form3.StringGrid1.Cells[5,6]:=FloatToStrF(m-l,fffixed,19,4);
    k:=(StrToFloat(Form1.StringGrid1.Cells[1,17])+StrToFloat(Form1.StringGrid1.Cells[1,8])+StrToFloat(Form1.StringGrid1.Cells[1,9])+StrToFloat(Form1.StringGrid1.Cells[1,10]))/StrToFloat(Form1.StringGrid1.Cells[1,16]);
    Form3.StringGrid1.Cells[1,7]:=FloatToStrF(k,fffixed,19,4);
    l:=(StrToFloat(Form1.StringGrid1.Cells[2,17])+StrToFloat(Form1.StringGrid1.Cells[2,8])+StrToFloat(Form1.StringGrid1.Cells[2,9])+StrToFloat(Form1.StringGrid1.Cells[2,10]))/StrToFloat(Form1.StringGrid1.Cells[2,16]);
    Form3.StringGrid1.Cells[2,7]:=FloatToStrF(l,fffixed,19,4);
    m:=(StrToFloat(Form1.StringGrid1.Cells[3,17])+StrToFloat(Form1.StringGrid1.Cells[3,8])+StrToFloat(Form1.StringGrid1.Cells[3,9])+StrToFloat(Form1.StringGrid1.Cells[3,10]))/StrToFloat(Form1.StringGrid1.Cells[3,16]);
    Form3.StringGrid1.Cells[3,7]:=FloatToStrF(m,fffixed,19,4);
    Form3.StringGrid1.Cells[4,7]:=FloatToStrF(l-k,fffixed,19,4);
    Form3.StringGrid1.Cells[5,7]:=FloatToStrF(m-l,fffixed,19,4);
 Form3.Show();
 Form1.Visible:=false;
end;

procedure TForm1.StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
StringGrid1.ColWidths[0] := 220;
end;

end.
