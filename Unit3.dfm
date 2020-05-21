object Form3: TForm3
  Left = 593
  Top = 185
  Width = 665
  Height = 585
  Caption = #1060#1080#1085#1072#1085#1089#1086#1074#1072#1103' '#1091#1089#1090#1086#1081#1095#1080#1074#1086#1089#1090#1100
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  DesignSize = (
    649
    526)
  PixelsPerInch = 96
  TextHeight = 13
  object BitBtn1: TBitBtn
    Left = 0
    Top = 477
    Width = 649
    Height = 49
    Anchors = [akLeft, akRight, akBottom]
    Caption = #1042#1099#1076#1072#1090#1100' '#1072#1085#1072#1083#1080#1079
    TabOrder = 0
    OnClick = BitBtn1Click
  end
  object StringGrid1: TStringGrid
    Left = 0
    Top = 0
    Width = 649
    Height = 478
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 1
  end
  object MainMenu1: TMainMenu
    Left = 432
    Top = 96
    object N1: TMenuItem
      Caption = #1057#1074#1103#1079#1100' '#1089' '#1085#1072#1084#1080
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = #1053#1072#1079#1072#1076
      OnClick = N2Click
    end
  end
end
