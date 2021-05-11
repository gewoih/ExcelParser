object Form2: TForm2
  Left = 0
  Top = 0
  Caption = 'Form2'
  ClientHeight = 399
  ClientWidth = 658
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 16
  object SuppliersTree: TVirtualStringTree
    Left = 0
    Top = 0
    Width = 658
    Height = 367
    AccessibleName = #1053#1072#1079#1074#1072#1085#1080#1077
    Align = alClient
    Colors.BorderColor = 15987699
    Colors.DisabledColor = clGray
    Colors.DropMarkColor = 15385233
    Colors.DropTargetColor = 15385233
    Colors.DropTargetBorderColor = 15987699
    Colors.FocusedSelectionColor = 15385233
    Colors.FocusedSelectionBorderColor = clWhite
    Colors.GridLineColor = 15987699
    Colors.HeaderHotColor = clBlack
    Colors.HotColor = clBlack
    Colors.SelectionRectangleBlendColor = 15385233
    Colors.SelectionRectangleBorderColor = 15385233
    Colors.SelectionTextColor = clBlack
    Colors.TreeLineColor = 9471874
    Colors.UnfocusedColor = 9694020
    Colors.UnfocusedSelectionColor = 15385233
    Colors.UnfocusedSelectionBorderColor = 15385233
    Colors.HeaderColor = 9694020
    Header.AutoSizeIndex = 2
    Header.Font.Charset = DEFAULT_CHARSET
    Header.Font.Color = clWindowText
    Header.Font.Height = -11
    Header.Font.Name = 'Tahoma'
    Header.Font.Style = []
    Header.Options = [hoAutoResize, hoColumnResize, hoDrag, hoShowSortGlyphs, hoVisible]
    TabOrder = 0
    OnBeforeCellPaint = SuppliersTreeBeforeCellPaint
    OnGetText = SuppliersTreeGetText
    Columns = <
      item
        Alignment = taCenter
        Position = 0
        Width = 80
        Aggregate = False
        FilterMode = 0
        WideText = 'id'
      end
      item
        Alignment = taCenter
        Position = 1
        Width = 150
        Aggregate = False
        FilterMode = 0
        WideText = #1048#1053#1053
      end
      item
        Position = 2
        Width = 424
        Aggregate = False
        FilterMode = 0
        WideText = #1053#1072#1079#1074#1072#1085#1080#1077
      end>
  end
  object btAddSupplier: TButton
    Left = 0
    Top = 367
    Width = 658
    Height = 32
    Align = alBottom
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100
    TabOrder = 1
    OnClick = btAddSupplierClick
  end
  object scGetSuppliersList: TStringContainer
    Items.Strings = (
      'select'
      '  sd.id,'
      '  VTK.dbo.AsStr(sd.id, '#39'2523'#39', 0),'
      '  VTK.dbo.AsStr(sd.id, '#39'2527'#39', 0)'
      'from VTK.dbo.spr_prop sp'
      'join VTK.dbo.spr_data sd on sd.id = sp.linkid and sd.mark=1'
      'where sp.metaid = 2326 and VTK.dbo.AsStr(sd.id, '#39'2527'#39', 0) > '#39#39
      'order by 3')
    Left = 40
    Top = 32
  end
end
