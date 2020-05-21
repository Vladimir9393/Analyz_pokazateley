object Form1: TForm1
  Left = 468
  Top = 159
  Width = 867
  Height = 655
  Caption = #1048#1089#1093#1086#1076#1085#1099#1077' '#1076#1072#1085#1085#1099#1077
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  DesignSize = (
    851
    596)
  PixelsPerInch = 96
  TextHeight = 13
  object StringGrid1: TStringGrid
    Left = 0
    Top = 0
    Width = 850
    Height = 500
    Anchors = [akLeft, akTop, akRight, akBottom]
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
    TabOrder = 0
    OnDrawCell = StringGrid1DrawCell
  end
  object Button1: TButton
    Left = 0
    Top = 499
    Width = 849
    Height = 49
    Anchors = [akLeft, akBottom]
    Caption = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 424
    Top = 547
    Width = 425
    Height = 49
    Anchors = [akLeft, akBottom]
    Caption = #1056#1072#1089#1095#1077#1090' '#1092#1080#1085'. '#1091#1089#1090#1081#1095#1080#1074#1086#1089#1090#1080
    TabOrder = 2
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 0
    Top = 547
    Width = 425
    Height = 49
    Anchors = [akLeft, akBottom]
    Caption = #1056#1072#1089#1095#1077#1090' '#1087#1083#1072#1090#1077#1078#1077#1089#1087#1086#1089#1086#1073#1085#1086#1089#1090#1080
    TabOrder = 3
    OnClick = Button3Click
  end
  object OpenDialog1: TOpenDialog
    Left = 656
    Top = 576
  end
  object MainMenu1: TMainMenu
    Left = 752
    Top = 424
    object N1: TMenuItem
      Caption = #1057#1074#1103#1079#1100' '#1089' '#1085#1072#1084#1080
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = #1042#1099#1093#1086#1076
      OnClick = N2Click
    end
  end
end
