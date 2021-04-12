object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 626
  ClientWidth = 1022
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 16
  object MainPanel: TPanel
    Left = 0
    Top = 0
    Width = 1022
    Height = 626
    Align = alClient
    TabOrder = 0
    object btExcelOpen: TButton
      Left = 1
      Top = 564
      Width = 1020
      Height = 61
      Align = alBottom
      Caption = #1042#1099#1073#1088#1072#1090#1100' '#1076#1086#1082#1091#1084#1077#1085#1090' Excel'
      TabOrder = 0
      OnClick = btExcelOpenClick
    end
    object SuppliersTree: TVirtualStringTree
      Left = 1
      Top = 1
      Width = 200
      Height = 563
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
      Colors.UnfocusedColor = 9693824
      Colors.UnfocusedSelectionColor = 15385233
      Colors.UnfocusedSelectionBorderColor = 15385233
      Colors.HeaderColor = 9693824
      DefaultNodeHeight = 25
      Header.AutoSizeIndex = 0
      Header.DefaultHeight = 25
      Header.Font.Charset = DEFAULT_CHARSET
      Header.Font.Color = clWindowText
      Header.Font.Height = -13
      Header.Font.Name = 'Tahoma'
      Header.Font.Style = []
      Header.Height = 26
      Header.Options = [hoAutoResize, hoColumnResize, hoDrag, hoShowSortGlyphs, hoVisible]
      PopupMenu = PopupMenu1
      TabOrder = 1
      TreeOptions.MiscOptions = [toAcceptOLEDrop, toCheckSupport, toEditable, toFullRepaintOnResize, toGridExtensions, toInitOnSave, toToggleOnDblClick, toWheelPanning, toVariableNodeHeight, toEditOnClick, toEditOnDblClick]
      TreeOptions.PaintOptions = [toHideFocusRect, toHideSelection, toShowDropmark, toShowHorzGridLines, toAlwaysHideSelection]
      ExplicitLeft = -5
      ExplicitTop = -5
      Columns = <
        item
          Alignment = taCenter
          CaptionAlignment = taCenter
          Options = [coAllowClick, coDraggable, coEnabled, coParentBidiMode, coParentColor, coResizable, coShowDropMark, coVisible, coAllowFocus, coUseCaptionAlignment, coEditable]
          Position = 0
          Width = 200
          Aggregate = False
          FilterMode = 0
          WideText = #1055#1086#1089#1090#1072#1074#1097#1080#1082
        end>
    end
    object FlatPanel1: TFlatPanel
      Tag = 1
      Left = 201
      Top = 1
      Width = 820
      Height = 563
      ParentColor = True
      ColorHighLight = 8623776
      ColorShadow = 8623776
      BorderColor = 8623776
      Align = alClient
      TabOrder = 2
      UseDockManager = True
      ExplicitLeft = 129
      ExplicitTop = -55
      object PreviewTree: TVirtualStringTree
        Left = 1
        Top = 1
        Width = 818
        Height = 561
        AccessibleName = #1054#1089#1090#1072#1090#1086#1082
        Align = alClient
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
        Colors.UnfocusedColor = 9693628
        Colors.UnfocusedSelectionColor = 15385233
        Colors.UnfocusedSelectionBorderColor = 15385233
        Colors.HeaderColor = 9693628
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
        Header.Options = [hoColumnResize, hoShowSortGlyphs, hoVisible]
        TabOrder = 0
        TreeOptions.AutoOptions = [toAutoDropExpand, toAutoScrollOnExpand, toAutoSort, toAutoSpanColumns, toAutoTristateTracking, toAutoDeleteMovedNodes, toAutoChangeScale]
        TreeOptions.MiscOptions = [toAcceptOLEDrop, toCheckSupport, toEditable, toFullRepaintOnResize, toGridExtensions, toInitOnSave, toToggleOnDblClick, toWheelPanning, toVariableNodeHeight, toEditOnClick, toEditOnDblClick]
        TreeOptions.PaintOptions = [toHideFocusRect, toHideSelection, toShowDropmark, toShowHorzGridLines, toAlwaysHideSelection]
        OnGetText = PreviewTreeGetText
        ExplicitHeight = 556
        Columns = <>
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = #1044#1086#1082#1091#1084#1077#1085#1090' Excel|*.xls|'#1058#1072#1073#1083#1080#1094#1072' Excel|*.csv'
    Left = 40
    Top = 40
  end
  object PopupMenu1: TPopupMenu
    Left = 120
    Top = 40
    object miAddSupplier: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1087#1086#1089#1090#1072#1074#1097#1080#1082#1072
    end
  end
end
