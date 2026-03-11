object Form1: TForm1
  Left = 729
  Top = 227
  Width = 474
  Height = 373
  Caption = 'Test'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    458
    335)
  PixelsPerInch = 96
  TextHeight = 13
  object btnProcess: TButton
    Left = 8
    Top = 8
    Width = 409
    Height = 25
    Anchors = [akLeft, akTop, akRight]
    Caption = 'Run'
    TabOrder = 0
    OnClick = btnProcessClick
  end
  object btnCancel: TButton
    Left = 348
    Top = 273
    Width = 73
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Cancel'
    Enabled = False
    TabOrder = 1
    OnClick = btnCancelClick
  end
  object MemoLog: TMemo
    Left = 8
    Top = 40
    Width = 409
    Height = 225
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 2
  end
  object ProgressBar1: TProgressBar
    Left = 8
    Top = 273
    Width = 337
    Height = 25
    Anchors = [akLeft, akRight, akBottom]
    TabOrder = 3
  end
  object OpenDialog1: TOpenDialog
    Left = 72
    Top = 72
  end
end
