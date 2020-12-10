object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 626
  ClientWidth = 1022
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object MainPanel: TPanel
    Left = 0
    Top = 0
    Width = 1022
    Height = 626
    Align = alClient
    TabOrder = 0
    object ChooseButton: TButton
      Left = 1
      Top = 564
      Width = 1020
      Height = 61
      Align = alBottom
      Caption = #1042#1099#1073#1088#1072#1090#1100' '#1076#1086#1082#1091#1084#1077#1085#1090' Excel'
      TabOrder = 0
      OnClick = ChooseButtonClick
    end
    object StringGrid1: TStringGrid
      Left = 1
      Top = 1
      Width = 1020
      Height = 563
      Align = alClient
      FixedCols = 0
      RowCount = 1
      FixedRows = 0
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = #1044#1086#1082#1091#1084#1077#1085#1090' Excel|*.xls'
    Left = 488
    Top = 352
  end
end
