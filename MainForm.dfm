object Form1: TForm1
  Left = 0
  Top = 0
  ActiveControl = PreviewTree
  Caption = 'ExcelParser'
  ClientHeight = 689
  ClientWidth = 1384
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 16
  object MainPanel: TPanel
    Left = 0
    Top = 0
    Width = 1384
    Height = 689
    Align = alClient
    TabOrder = 0
    object SuppliersTree: TVirtualStringTree
      Left = 1
      Top = 1
      Width = 360
      Height = 687
      AccessibleName = #1054#1089#1090#1072#1090#1086#1082
      Align = alLeft
      BorderStyle = bsNone
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
      Colors.UnfocusedColor = 9687932
      Colors.UnfocusedSelectionColor = 15385233
      Colors.UnfocusedSelectionBorderColor = 15385233
      Colors.HeaderColor = 9687932
      DefaultNodeHeight = 25
      Header.AutoSizeIndex = 3
      Header.DefaultHeight = 25
      Header.Font.Charset = DEFAULT_CHARSET
      Header.Font.Color = clWindowText
      Header.Font.Height = -13
      Header.Font.Name = 'Tahoma'
      Header.Font.Style = []
      Header.Height = 28
      Header.Options = [hoAutoResize, hoColumnResize, hoDrag, hoShowSortGlyphs, hoVisible]
      PopupMenu = SuppliersTreePopupMenu
      TabOrder = 0
      TreeOptions.MiscOptions = [toAcceptOLEDrop, toCheckSupport, toEditable, toFullRepaintOnResize, toGridExtensions, toInitOnSave, toToggleOnDblClick, toWheelPanning, toVariableNodeHeight, toEditOnClick, toEditOnDblClick]
      TreeOptions.PaintOptions = [toHideFocusRect, toHideSelection, toShowDropmark, toShowHorzGridLines, toAlwaysHideSelection]
      OnBeforeCellPaint = SuppliersTreeBeforeCellPaint
      OnFocusChanged = SuppliersTreeFocusChanged
      OnGetText = SuppliersTreeGetText
      Columns = <
        item
          Alignment = taCenter
          CaptionAlignment = taCenter
          Options = [coAllowClick, coDraggable, coEnabled, coParentBidiMode, coParentColor, coResizable, coShowDropMark, coVisible, coAllowFocus, coUseCaptionAlignment, coEditable]
          Position = 0
          Width = 40
          Aggregate = False
          FilterMode = 0
          WideText = 'id'
        end
        item
          Alignment = taCenter
          Position = 1
          Width = 80
          Aggregate = False
          FilterMode = 0
          WideText = 'pid'
        end
        item
          Alignment = taCenter
          Position = 2
          Width = 100
          Aggregate = False
          FilterMode = 0
          WideText = #1048#1053#1053
        end
        item
          Alignment = taCenter
          Position = 3
          Width = 140
          Aggregate = False
          FilterMode = 0
          WideText = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
        end>
    end
    object FlatPanel1: TFlatPanel
      Tag = 1
      Left = 361
      Top = 1
      Width = 1022
      Height = 687
      ParentColor = True
      ColorHighLight = 8623776
      ColorShadow = 8623776
      BorderColor = 8623776
      Align = alClient
      TabOrder = 1
      UseDockManager = True
      object Splitter1: TSplitter
        Left = 1
        Top = 394
        Width = 1020
        Height = 3
        Cursor = crVSplit
        Align = alTop
        ExplicitTop = 369
        ExplicitWidth = 317
      end
      object ExcelTree: TVirtualStringTree
        Left = 1
        Top = 26
        Width = 1020
        Height = 368
        AccessibleName = #1054#1089#1090#1072#1090#1086#1082
        Align = alTop
        BorderStyle = bsNone
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
        Colors.UnfocusedColor = 9687736
        Colors.UnfocusedSelectionColor = 15385233
        Colors.UnfocusedSelectionBorderColor = 15385233
        Colors.HeaderColor = 9687736
        DefaultNodeHeight = 25
        Header.AutoSizeIndex = -1
        Header.DefaultHeight = 25
        Header.Font.Charset = DEFAULT_CHARSET
        Header.Font.Color = clWindowText
        Header.Font.Height = -13
        Header.Font.Name = 'Tahoma'
        Header.Font.Style = []
        Header.Height = 26
        Header.MainColumn = -1
        Margin = 8
        TabOrder = 0
        TreeOptions.AutoOptions = [toAutoDropExpand, toAutoScrollOnExpand, toAutoSort, toAutoSpanColumns, toAutoTristateTracking, toAutoDeleteMovedNodes, toAutoChangeScale]
        TreeOptions.PaintOptions = [toHideFocusRect, toHideSelection, toShowDropmark, toShowHorzGridLines, toAlwaysHideSelection]
        OnBeforeCellPaint = SuppliersTreeBeforeCellPaint
        OnGetText = ExcelTreeGetText
        OnKeyPress = ExcelTreeKeyPress
        Columns = <>
      end
      object FlatPanel2: TFlatPanel
        Tag = 1
        Left = 1
        Top = 397
        Width = 1020
        Height = 289
        ParentColor = True
        ColorHighLight = 8623776
        ColorShadow = 8623776
        BorderColor = 8623776
        Align = alClient
        TabOrder = 1
        UseDockManager = True
        object PreviewTree: TVirtualStringTree
          Left = 1
          Top = 121
          Width = 1018
          Height = 118
          AccessibleName = #1054#1089#1090#1072#1090#1086#1082
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
          Colors.UnfocusedColor = 9687540
          Colors.UnfocusedSelectionColor = 15385233
          Colors.UnfocusedSelectionBorderColor = 15385233
          Colors.HeaderColor = 9687540
          DefaultNodeHeight = 25
          Header.AutoSizeIndex = 0
          Header.DefaultHeight = 25
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -13
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          Header.Height = 26
          TabOrder = 0
          OnGetText = PreviewTreeGetText
          Columns = <
            item
              CaptionAlignment = taCenter
              Options = [coAllowClick, coDraggable, coEnabled, coParentBidiMode, coParentColor, coResizable, coShowDropMark, coVisible, coAllowFocus, coUseCaptionAlignment, coEditable]
              Position = 0
              Width = 200
              Aggregate = False
              FilterMode = 0
              WideText = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
            end
            item
              CaptionAlignment = taCenter
              Options = [coAllowClick, coDraggable, coEnabled, coParentBidiMode, coParentColor, coResizable, coShowDropMark, coVisible, coAllowFocus, coUseCaptionAlignment, coEditable]
              Position = 1
              Width = 120
              Aggregate = False
              FilterMode = 0
              WideText = #1040#1088#1090#1080#1082#1091#1083
            end
            item
              CaptionAlignment = taCenter
              Options = [coAllowClick, coDraggable, coEnabled, coParentBidiMode, coParentColor, coResizable, coShowDropMark, coVisible, coAllowFocus, coUseCaptionAlignment, coEditable]
              Position = 2
              Width = 100
              Aggregate = False
              FilterMode = 0
              WideText = #1041#1072#1079#1086#1074#1072#1103' '#1094#1077#1085#1072
            end
            item
              CaptionAlignment = taCenter
              Options = [coAllowClick, coDraggable, coEnabled, coParentBidiMode, coParentColor, coResizable, coShowDropMark, coVisible, coAllowFocus, coUseCaptionAlignment, coEditable]
              Position = 3
              Width = 120
              Aggregate = False
              FilterMode = 0
              WideText = #1062#1077#1085#1072' '#1074#1093#1086#1076' '#1042#1058#1050
            end
            item
              CaptionAlignment = taCenter
              Options = [coAllowClick, coDraggable, coEnabled, coParentBidiMode, coParentColor, coResizable, coShowDropMark, coVisible, coAllowFocus, coUseCaptionAlignment, coEditable]
              Position = 4
              Width = 100
              Aggregate = False
              FilterMode = 0
              WideText = #1054#1089#1090#1072#1090#1086#1082
            end
            item
              CaptionAlignment = taCenter
              Options = [coAllowClick, coDraggable, coEnabled, coParentBidiMode, coParentColor, coResizable, coShowDropMark, coVisible, coAllowFocus, coUseCaptionAlignment, coEditable]
              Position = 5
              Width = 100
              Aggregate = False
              FilterMode = 0
              WideText = #1043#1086#1076
            end
            item
              Position = 6
              Aggregate = False
              FilterMode = 0
              WideText = #1054#1073#1098#1077#1084
            end
            item
              Position = 7
              Aggregate = False
              FilterMode = 0
              WideText = #1050#1088#1077#1087#1086#1089#1090#1100
            end>
        end
        object FlatPanel3: TFlatPanel
          Tag = 1
          Left = 1
          Top = 1
          Width = 1018
          Height = 120
          ParentColor = True
          ColorHighLight = 8623776
          ColorShadow = 8623776
          BorderColor = 8623776
          Align = alTop
          TabOrder = 1
          UseDockManager = True
          object Label3: TLabel
            Left = 314
            Top = 17
            Width = 107
            Height = 16
            Caption = #1041#1072#1079#1086#1074#1072#1103' '#1094#1077#1085#1072' (3):'
          end
          object Label1: TLabel
            Left = 568
            Top = 17
            Width = 73
            Height = 16
            Caption = #1054#1089#1090#1072#1090#1082#1080' (5):'
          end
          object Label2: TLabel
            Left = 21
            Top = 17
            Width = 132
            Height = 16
            Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1040#1055' (1):'
          end
          object Label4: TLabel
            Left = 60
            Top = 44
            Width = 93
            Height = 16
            Caption = #1040#1088#1090#1080#1082#1091#1083' '#1040#1055' (2):'
          end
          object Label5: TLabel
            Left = 594
            Top = 49
            Width = 47
            Height = 16
            Caption = #1043#1086#1076' (6):'
          end
          object Label7: TLabel
            Left = 307
            Top = 49
            Width = 113
            Height = 16
            Caption = #1062#1077#1085#1072' '#1074#1093#1086#1076' '#1042#1058#1050' (4):'
          end
          object Label8: TLabel
            Left = 798
            Top = 48
            Width = 80
            Height = 16
            Caption = #1050#1088#1077#1087#1086#1089#1090#1100' (8):'
          end
          object Label9: TLabel
            Left = 814
            Top = 17
            Width = 64
            Height = 16
            Caption = #1054#1073#1098#1077#1084' (7):'
          end
          object btSaveLinks: TButton
            Left = 1
            Top = 86
            Width = 1016
            Height = 33
            Align = alBottom
            Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1087#1088#1080#1074#1103#1079#1082#1091
            TabOrder = 0
            OnClick = btSaveLinksClick
          end
          object tName: TEdit
            Left = 159
            Top = 14
            Width = 121
            Height = 24
            Enabled = False
            TabOrder = 1
          end
          object tArticle: TEdit
            Left = 159
            Top = 44
            Width = 121
            Height = 24
            Enabled = False
            TabOrder = 2
          end
          object tBasePrice: TEdit
            Left = 426
            Top = 14
            Width = 121
            Height = 24
            Enabled = False
            TabOrder = 3
          end
          object tQuantity: TEdit
            Left = 647
            Top = 14
            Width = 121
            Height = 24
            Enabled = False
            TabOrder = 4
          end
          object tYear: TEdit
            Left = 647
            Top = 44
            Width = 121
            Height = 24
            Enabled = False
            TabOrder = 5
          end
          object tPriceIn: TEdit
            Left = 426
            Top = 44
            Width = 121
            Height = 24
            Enabled = False
            TabOrder = 6
          end
        end
        object FlatPanel4: TFlatPanel
          Tag = 1
          Left = 1
          Top = 239
          Width = 1018
          Height = 49
          ParentColor = True
          ColorHighLight = 8623776
          ColorShadow = 8623776
          BorderColor = 8623776
          Align = alBottom
          TabOrder = 2
          UseDockManager = True
          object label6: TLabel
            Left = 24
            Top = 16
            Width = 120
            Height = 16
            BiDiMode = bdLeftToRight
            Caption = #1050#1086#1076' '#1087#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103':'
            ParentBiDiMode = False
          end
          object btUploadPrice: TButton
            AlignWithMargins = True
            Left = 823
            Top = 4
            Width = 191
            Height = 41
            Align = alRight
            Caption = #1054#1090#1087#1088#1072#1074#1080#1090#1100' '#1087#1088#1072#1081#1089
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 0
            OnClick = btUploadPriceClick
          end
          object cbDivisions: TComboBox
            Left = 159
            Top = 13
            Width = 270
            Height = 24
            TabOrder = 1
          end
        end
      end
      object tbExcelTabs: TTabSet
        Left = 1
        Top = 1
        Width = 1020
        Height = 25
        Align = alTop
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        OnChange = tbExcelTabsChange
      end
    end
  end
  object tVolume: TEdit
    Left = 1247
    Top = 413
    Width = 121
    Height = 24
    Enabled = False
    TabOrder = 1
  end
  object tStrength: TEdit
    Left = 1247
    Top = 443
    Width = 121
    Height = 24
    Enabled = False
    TabOrder = 2
  end
  object OpenDialog1: TOpenDialog
    Filter = #1044#1086#1082#1091#1084#1077#1085#1090#1099' '#1080' '#1090#1072#1073#1083#1080#1094#1099' Excel|*.xls; *.csv; *.xlsx'
    Left = 40
    Top = 40
  end
  object SuppliersTreePopupMenu: TPopupMenu
    Left = 144
    Top = 40
    object miAddSupplier: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1087#1086#1089#1090#1072#1074#1097#1080#1082#1072
      OnClick = miAddSupplierClick
    end
    object miAddExcel: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1087#1088#1072#1081#1089
      OnClick = miAddExcelClick
    end
    object miDeleteSupplier: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100' '#1087#1086#1089#1090#1072#1074#1097#1080#1082#1072
      OnClick = miDeleteSupplierClick
    end
  end
  object scLoadSuppliers: TStringContainer
    Items.Strings = (
      'select'
      '  VTK_Excel.dbo.Suppliers.id,'
      '  sd.id as pid,'
      '  VTK.dbo.AsStr(sd.id, '#39'2523'#39', 0) as tin,'
      '  VTK.dbo.AsStr(sd.id, '#39'2527'#39', 0) as name'
      'from VTK.dbo.spr_prop sp'
      'join VTK.dbo.spr_data sd on sd.id = sp.linkid and sd.mark=1'
      'join VTK_Excel.dbo.Suppliers on supplier_id = sd.id'
      'where sp.metaid = 2326 and VTK.dbo.AsStr(sd.id, '#39'2527'#39', 0) > '#39#39
      'order by 1'
      '')
    Left = 40
    Top = 104
  end
  object scLoadExcel: TStringContainer
    Items.Strings = (
      
        'declare @rmax int, @cmax int, @rmin int, @cmin int, @linkid int,' +
        ' @c varchar(max)'
      ''
      'select'
      '  @linkid = 1,'
      '  @rmin = 0,'
      '  @cmin = 0,'
      '  @rmax = max(row),'
      '  @cmax = max(col)'
      'from Excel_templates where val is not null'
      ''
      ';with r as'
      '('
      'select @cmin a'
      'union all'
      'select a + 1 from r where a < @cmax'
      ')'
      
        'select @c = STRING_AGG('#39'['#39' + cast(a as varchar) +'#39']'#39', '#39','#39') from ' +
        'r'
      ''
      'declare @s varchar(max) = '#39
      'select * from'
      '('
      
        'select row, col, val from Excel_templates where val is not null ' +
        'and row > 0 and linkid = %d'
      ') a'
      'pivot (max(val) for col in ('#39' + @c + '#39')) p'
      'order by 1'
      #39
      ''
      'exec(@s)')
    Left = 120
    Top = 104
  end
  object scAddSupplier: TStringContainer
    Items.Strings = (
      'insert into Suppliers values(%s)'
      'declare @id int = SCOPE_IDENTITY()'
      'insert into Excel_templates values'
      '(@id, 0, 1, '#39'-1'#39'),'
      '(@id, 0, 2, '#39'-1'#39'),'
      '(@id, 0, 3, '#39'-1'#39'),'
      '(@id, 0, 4, '#39'-1'#39'),'
      '(@id, 0, 5, '#39'-1'#39'),'
      '(@id, 0, 6, '#39'-1'#39'),'
      '(@id, 0, 7, '#39'-1'#39'),'
      '(@id, 0, 8, '#39'-1'#39')')
    Left = 40
    Top = 168
  end
  object scLoadDivisions: TStringContainer
    Items.Strings = (
      
        'select linkid, VTK.dbo.AsStr(linkid, 2597, 12) as code,  VTK.dbo' +
        '.AsStr(linkid, 4406, 0) as name'
      'from VTK.dbo.spr_prop sp'
      'join  VTK.dbo.spr_data sd on sd.id = sp.linkid and sd.mark = 1'
      'where sp.metaid = 4340 and sp.value = %d order by name')
    Left = 120
    Top = 168
  end
end
