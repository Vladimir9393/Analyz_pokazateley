object Form1: TForm1
  Left = 481
  Top = 159
  Width = 866
  Height = 672
  Caption = #1048#1089#1093#1086#1076#1085#1099#1077' '#1076#1072#1085#1085#1099#1077
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object StringGrid1: TStringGrid
    Left = 0
    Top = 0
    Width = 849
    Height = 537
    Anchors = [akBottom,akTop,akLeft,akRight]
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
    TabOrder = 0
    OnDrawCell = StringGrid1DrawCell
  end
  object Button1: TButton
    Left = 0
    Top = 536
    Width = 849
    Height = 49
    Anchors = [akBottom,akLeft]
    Caption = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 296
    Top = 584
    Width = 297
    Height = 49
    Anchors = [akBottom,akLeft]
    Caption = #1056#1072#1089#1095#1077#1090' '#1092#1080#1085'. '#1091#1089#1090#1081#1095#1080#1074#1086#1089#1090#1080
    TabOrder = 2
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 0
    Top = 584
    Width = 297
    Height = 49
    Anchors = [akBottom,akLeft]
    Caption = #1056#1072#1089#1095#1077#1090' '#1087#1083#1072#1090#1077#1078#1077#1089#1087#1086#1089#1086#1073#1085#1086#1089#1090#1080
    TabOrder = 3
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 592
    Top = 584
    Width = 257
    Height = 49
    Anchors = [akBottom,akLeft]
    Caption = #1042#1099#1093#1086#1076
    TabOrder = 4
    OnClick = Button4Click
  end
  object OpenDialog1: TOpenDialog
    Left = 656
    Top = 576
  end
end
