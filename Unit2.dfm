object Form2: TForm2
  Left = 569
  Top = 188
  Width = 649
  Height = 568
  Caption = #1055#1083#1072#1090#1077#1078#1077#1089#1087#1086#1089#1086#1073#1085#1086#1089#1090#1100
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  DesignSize = (
    633
    509)
  PixelsPerInch = 96
  TextHeight = 13
  object StringGrid: TStringGrid
    Left = 0
    Top = 0
    Width = 633
    Height = 461
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 0
  end
  object BitBtn1: TBitBtn
    Left = 0
    Top = 460
    Width = 633
    Height = 49
    Anchors = [akLeft, akRight, akBottom]
    Caption = #1042#1099#1076#1072#1090#1100' '#1072#1085#1072#1083#1080#1079
    TabOrder = 1
    OnClick = BitBtn1Click
  end
  object MainMenu1: TMainMenu
    Left = 560
    Top = 304
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
