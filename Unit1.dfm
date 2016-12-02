object Form1: TForm1
  Left = 238
  Top = 214
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Form1'
  ClientHeight = 395
  ClientWidth = 1025
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 344
    Top = 96
    Width = 342
    Height = 58
    Caption = #1042#1080#1073#1077#1088#1110#1090#1100' '#1092#1072#1081#1083
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -48
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Button1: TButton
    Left = 16
    Top = 288
    Width = 75
    Height = 25
    Caption = #1042#1110#1076#1082#1088#1080#1090#1080
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 104
    Top = 288
    Width = 75
    Height = 25
    Caption = #1042#1080#1093#1110#1076
    TabOrder = 1
    OnClick = Button2Click
  end
  object Chart1: TChart
    Left = 0
    Top = 0
    Width = 1025
    Height = 273
    AllowPanning = pmNone
    AllowZoom = False
    BackWall.Brush.Color = clWhite
    BackWall.Brush.Style = bsClear
    MarginBottom = 5
    MarginTop = 5
    Title.Text.Strings = (
      'TChart')
    Title.Visible = False
    BottomAxis.Automatic = False
    BottomAxis.AutomaticMaximum = False
    BottomAxis.AutomaticMinimum = False
    LeftAxis.Automatic = False
    LeftAxis.AutomaticMinimum = False
    Legend.Visible = False
    RightAxis.Visible = False
    TopAxis.Visible = False
    View3D = False
    TabOrder = 2
    Visible = False
    object Series1: TLineSeries
      Marks.ArrowLength = 8
      Marks.Font.Charset = DEFAULT_CHARSET
      Marks.Font.Color = clBlack
      Marks.Font.Height = -8
      Marks.Font.Name = 'Arial'
      Marks.Font.Style = []
      Marks.Visible = False
      SeriesColor = clRed
      Pointer.InflateMargins = True
      Pointer.Style = psRectangle
      Pointer.Visible = False
      XValues.DateTime = False
      XValues.Name = 'X'
      XValues.Multiplier = 1
      XValues.Order = loAscending
      YValues.DateTime = False
      YValues.Name = 'Y'
      YValues.Multiplier = 1
      YValues.Order = loNone
    end
  end
  object RadioGroup1: TRadioGroup
    Left = 200
    Top = 288
    Width = 185
    Height = 105
    Caption = #1042#1080#1073#1110#1088' '#1090#1080#1087#1091' '#1075#1088#1072#1092#1110#1082#1072
    Enabled = False
    ItemIndex = 0
    Items.Strings = (
      #1047#1072#1075#1072#1083#1100#1085#1080#1081
      #1052#1110#1089#1103#1094#1110
      #1044#1085#1110)
    TabOrder = 3
    OnClick = RadioGroup1Click
  end
  object ComboBox1: TComboBox
    Left = 16
    Top = 328
    Width = 169
    Height = 21
    DropDownCount = 12
    Enabled = False
    ItemHeight = 13
    ItemIndex = 0
    TabOrder = 4
    Text = #1055#1086' '#1084#1110#1089#1103#1094#1103#1093
    OnChange = ComboBox1Change
    Items.Strings = (
      #1055#1086' '#1084#1110#1089#1103#1094#1103#1093
      #1055#1086' '#1076#1085#1103#1093)
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Excel|*.xls;*.xlsx'
    Left = 32
    Top = 360
  end
  object MainMenu1: TMainMenu
    Left = 72
    Top = 360
    object N1: TMenuItem
      Caption = #1060#1072#1081#1083
      object N2: TMenuItem
        Caption = #1042#1110#1076#1082#1088#1080#1090#1080
        OnClick = N2Click
      end
      object N3: TMenuItem
        Caption = #1042#1080#1093#1110#1076
        OnClick = N3Click
      end
    end
    object N4: TMenuItem
      Caption = #1055#1088#1086' '#1087#1088#1086#1075#1088#1072#1084#1091
      OnClick = N4Click
    end
  end
end
