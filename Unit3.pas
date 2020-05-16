unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Grids, ComObj;

type
  TForm3 = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    StringGrid1: TStringGrid;
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

uses ProgramAnalys;

{$R *.dfm}

procedure TForm3.BitBtn2Click(Sender: TObject);
begin
 Form3.Close;
 Form1.Show();
end;

procedure TForm3.BitBtn1Click(Sender: TObject);
Var w:variant;
    str:String;
    a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u:double;
begin
 w:=CreateOleObject('Word.Application');
 w.Visible:=True;
 w.Documents.Add('E:\Образование\УНИВЕР\1. Дисциплины\ПС\ПС_МельникВВ\Вывод.docx');
 w.Selection.Start:=0;
 w.Selection.End:=0;
 w.Selection.Find.Forward:=True;
 a:=StrToFloat(StringGrid1.Cells[1,1]);
 b:=StrToFloat(StringGrid1.Cells[2,1]);
 c:=StrToFloat(StringGrid1.Cells[3,1]);
 d:=StrToFloat(StringGrid1.Cells[1,2]);
 e:=StrToFloat(StringGrid1.Cells[2,2]);
 f:=StrToFloat(StringGrid1.Cells[3,2]);
 g:=StrToFloat(StringGrid1.Cells[1,3]);
 h:=StrToFloat(StringGrid1.Cells[2,3]);
 i:=StrToFloat(StringGrid1.Cells[3,3]);
 j:=StrToFloat(StringGrid1.Cells[1,4]);
 k:=StrToFloat(StringGrid1.Cells[2,4]);
 l:=StrToFloat(StringGrid1.Cells[3,4]);
 m:=StrToFloat(StringGrid1.Cells[1,5]);
 n:=StrToFloat(StringGrid1.Cells[2,5]);
 o:=StrToFloat(StringGrid1.Cells[3,5]);
 p:=StrToFloat(StringGrid1.Cells[1,6]);
 q:=StrToFloat(StringGrid1.Cells[2,6]);
 r:=StrToFloat(StringGrid1.Cells[3,6]);
 s:=StrToFloat(StringGrid1.Cells[1,7]);
 t:=StrToFloat(StringGrid1.Cells[2,7]);
 u:=StrToFloat(StringGrid1.Cells[3,7]);
 if (((a>b) and (b<c) and (c>0.5)) and ((d<e) and (e>f) and (f<0.5)) and ((g>h) and (h<i) and (i>0.6)) and ((j<k) and (k>l) and (l<1)) and ((m>n) and (n>o)) and ((p<q) and (q<r) and (r>0.4)) and ((s>t) and (t<u)))
                                                                                   then str:='Исходя из расчетов, видно, что по сравнению с 2016 годом коэффициент автономности '+
                                                                                             ' увеличился, а коэффициент зависимости уменьшился в 2018-м году. '+
                                                                                             'Но все равно они сохраняют хорошие показатели и говорят о стабильности и независимости от внешних кредиторов. '+
                                                                                             'Незначительное изменение коэффициента финансовой устойчивости '+
                                                                                             'и изменение коэффициента финансовой активности '+
                                                                                             'не меняют общей картины и говорит о хорошей финансовой устойчивости '+
                                                                                             'предприятия и независимости от внешних инвесторов. '+
                                                                                             'Уменьшение коэффициента долгосрочного привлечения говорит о '+
                                                                                             'положительной тенденции к росту независимоти от инвесторов. '+
                                                                                             'Увеличение коэффициента мобильности собственных оборотных средств'+
                                                                                             ' показывает, что предприятие улучшило свое финансовое положение. '+
                                                                                             'Судя по снижению уровня коэффициента имущества производственного назначения,'+
                                                                                             ' относительно 2016 года, организации целесообразно привлечь долгосрочные заемные средства '+
                                                                                             'для увеличения имущества производственного назначения.'
 else
 if (((a<b) and (b<c) and (c>0.5)) and ((d<e) and (e<f) and (f<0.5)) and ((g<h) and (h<i) and (i>0.6)) and ((j<k) and (k<l) and (l<1)) and ((m<n) and (n<o)) and ((p<q) and (q<r) and (r>0.4)) and ((s<t) and (t<u)))
                                                                                   then str:='Исходя из расчетов, видно, что коэффициент автономности'+
                                                                                             ' уменьшается во всех трех периодах, также как и коэффициент зависимости. '+
                                                                                             'Но все равно они сохраняют хорошие показатели и говорят о стабильности и независимости от внешних кредиторов. '+
                                                                                             'Незначительное уменьшение коэффициента финансовой устойчивости '+
                                                                                             'и изменение коэффициента финансовой активности '+
                                                                                             'не меняют общей картины и говорит о хорошей финансовой устойчивости '+
                                                                                             'предприятия и независимости от внешних инвесторов. '+
                                                                                             'Уменьшение коэффициента долгосрочного привлечения говорит о '+
                                                                                             'положительной тенденции к росту независимоти от инвесторов. '+
                                                                                             'Уменьшение коэффициента мобильности собственных оборотных средств'+
                                                                                             ' указывает на то, что предприятие ухудшило свое финансовое положение. '+
                                                                                             'Судя по снижению уровня коэффициента имущества производственного назначения,'+
                                                                                             ' организации целесообразно привлечь долгосрочные заемные средства '+
                                                                                             'для увеличения имущества производственного назначения.';

 w.ActiveDocument.Range.Text:=str;
end;

end.
