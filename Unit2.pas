unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Grids, ComObj;

type
  TForm2 = class(TForm)
    StringGrid: TStringGrid;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

uses ProgramAnalys;

{$R *.dfm}

procedure TForm2.BitBtn2Click(Sender: TObject);
begin
 Form2.Close;
 Form1.Show();
end;

procedure TForm2.BitBtn1Click(Sender: TObject);
Var w:variant;
    str:String;
    a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p:double;
begin
 w:=CreateOleObject('Word.Application');
 w.Visible:=True;
 w.Documents.Add('E:\�����������\������\1. ����������\��\��_���������\�����.docx');
 w.Selection.Start:=0;
 w.Selection.End:=0;
 w.Selection.Find.Forward:=True;
 a:=StrToFloat(StringGrid.Cells[1,1]);
 b:=StrToFloat(StringGrid.Cells[2,1]);
 c:=StrToFloat(StringGrid.Cells[3,1]);
 d:=StrToFloat(StringGrid.Cells[1,2]);
 e:=StrToFloat(StringGrid.Cells[2,2]);
 f:=StrToFloat(StringGrid.Cells[3,2]);
 g:=StrToFloat(StringGrid.Cells[1,3]);
 h:=StrToFloat(StringGrid.Cells[2,3]);
 i:=StrToFloat(StringGrid.Cells[3,3]);
 j:=StrToFloat(StringGrid.Cells[1,4]);
 k:=StrToFloat(StringGrid.Cells[2,4]);
 l:=StrToFloat(StringGrid.Cells[3,4]);
 m:=StrToFloat(StringGrid.Cells[2,5]);
 n:=StrToFloat(StringGrid.Cells[3,5]);
 if (((a<b)and(b<c)) and ((d<e) and (e<f)) and ((g>h) and (h<i)) and ((j>k) and (k<l)) and (m>n)) then str:='������ �� ���������� ������, ����� ������� �����, ��� � 2017 ���� ��������� ���������'+
                                                                                                       ' ������� ����������� � �������������� ������������ ���������� ����������. '+
                                                                                                       '�� 2018 ��� ��������� ��������� ����������� �� ��������� � 2017 � 2016 �����. '+
                                                                                                       '� �����, ���������� ����� ��������� � �����. '+
                                                                                                       '������ ��������� ����������� �������������� ������������������ �����������, '+
                                                                                                       '���������� �� ����������� �������������� ���������� ������� ����������� �����������. '+
                                                                                                       '��� ����� ��������� ��������� �� ������������� ��������� �����������.'+
                                                                                                       ' ����������� ������� �������� ���������� ��������. '+
                                                                                                       '� ��������� ������ �������� ���������� ������.'
 else
 if (((a>b)and(b>c)) and ((d>e) and (e>f)) and ((g>h) and (h>i)) and ((j>k) and (k>l)) and (m>n)) then str:='������ �� ���������� ������, ����� ������� �����, �� ����������� ������ ��������� ���������'+
                                                                                                       ' ���� �����������, ������������ � �������. '+
                                                                                                       '���������� ����� ��������� � �������. '+
                                                                                                       '����� ��������� ����������� �������������� ������������������ �����������, '+
                                                                                                       '���������� �� ����������� �������������� ���������� ������� ����������� �����������. '+
                                                                                                       '��� ����� ��������� ��������� �� ������������� ��������� �����������. '+
                                                                                                       '����������� ������� ������������ ������������� ��������.'+
                                                                                                       ' � ��������� ������ �������� �������� �����������'
 else
 if (((a>b)and(b>c)) and ((d>e) and (e>f)) and ((g>h) and (h>i)) and ((j>k) and (k>l)) and (m<n)) then str:='������ �� ���������� ������, ����� ������� �����, ��� �� ���� ����������� ���� ����'+
                                                                                                       '������� � 2016 ���� ���� ����������� ���������� �������������. '+
                                                                                                       '���������� ����� ��������� � �������. '+
                                                                                                       '������ ��������� ����������� �������������� ������������������ �����������, '+
                                                                                                       '���������� �� ����������� �������������� ���������� ������� ����������� �����������. '+
                                                                                                       '� ����� ����������� ������� ������������ ���������� ������������� ��������. '+
                                                                                                       '��������� � ��������� ����� ��� ����� ������� �� ����� ��������������� ������������� �����������, '+
                                                                                                       '��� ����� ������� ������ �������� � ���������� �����������.';

 w.ActiveDocument.Range.Text:=str;
 end;

end.
