object frmGerarPlanilhaPayback: TfrmGerarPlanilhaPayback
  Left = 0
  Top = 0
  Caption = 'Gerar Planilha Payback'
  ClientHeight = 300
  ClientWidth = 437
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  TextHeight = 15
  object btnBrowse: TSpeedButton
    Left = 399
    Top = 37
    Width = 23
    Height = 22
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00555555555555
      55555555FFFFFFFFFF55555000000000055555577777777775F55500B8B8B8B8
      B05555775F555555575F550F0B8B8B8B8B05557F75F555555575550BF0B8B8B8
      B8B0557F575FFFFFFFF7550FBF0000000000557F557777777777500BFBFBFBFB
      0555577F555555557F550B0FBFBFBFBF05557F7F555555FF75550F0BFBFBF000
      55557F75F555577755550BF0BFBF0B0555557F575FFF757F55550FB700007F05
      55557F557777557F55550BFBFBFBFB0555557F555555557F55550FBFBFBFBF05
      55557FFFFFFFFF7555550000000000555555777777777755555550FBFB055555
      5555575FFF755555555557000075555555555577775555555555}
    NumGlyphs = 2
    OnClick = btnBrowseClick
  end
  object Label1: TLabel
    Left = 13
    Top = 188
    Width = 88
    Height = 15
    Caption = 'Produ'#231#227'o Anual:'
  end
  object Label2: TLabel
    Left = 156
    Top = 188
    Width = 89
    Height = 15
    Caption = 'Consumo Anual:'
  end
  object Label3: TLabel
    Left = 13
    Top = 12
    Width = 72
    Height = 15
    Caption = 'Planilha Base:'
  end
  object Label4: TLabel
    Left = 301
    Top = 209
    Width = 34
    Height = 15
    Caption = 'Label4'
  end
  object Label5: TLabel
    Left = 301
    Top = 188
    Width = 106
    Height = 15
    Caption = 'Investimento Inicial:'
  end
  object Label6: TLabel
    Left = 13
    Top = 76
    Width = 40
    Height = 15
    Caption = 'Cidade:'
  end
  object Label7: TLabel
    Left = 156
    Top = 76
    Width = 45
    Height = 15
    Caption = 'Arquivo:'
  end
  object Label8: TLabel
    Left = 301
    Top = 76
    Width = 50
    Height = 15
    Caption = 'M'#243'dulos:'
  end
  object Label9: TLabel
    Left = 13
    Top = 132
    Width = 28
    Height = 15
    Caption = 'KWp:'
  end
  object Label10: TLabel
    Left = 156
    Top = 132
    Width = 55
    Height = 15
    Caption = #193'rea (m2):'
  end
  object Button1: TButton
    Left = 170
    Top = 261
    Width = 89
    Height = 25
    Caption = 'Gerar Planilha'
    TabOrder = 0
    OnClick = Button1Click
  end
  object edtPlanilhaBase: TEdit
    Left = 13
    Top = 36
    Width = 380
    Height = 23
    Hint = 'Planilha Base'
    TabOrder = 1
  end
  object edtPdcAnual: TEdit
    Left = 13
    Top = 208
    Width = 121
    Height = 23
    Hint = 'Produ'#231#227'o Anual'
    TabOrder = 2
  end
  object edtConsAnual: TEdit
    Left = 156
    Top = 208
    Width = 121
    Height = 23
    Hint = 'Consumo Anual'
    TabOrder = 3
  end
  object edtInvestInic: TEdit
    Left = 301
    Top = 208
    Width = 121
    Height = 23
    Hint = 'Investimento Inicial'
    TabOrder = 4
  end
  object edtCidade: TEdit
    Left = 13
    Top = 96
    Width = 121
    Height = 23
    Hint = 'Cidade'
    TabOrder = 5
  end
  object edtArquivo: TEdit
    Left = 156
    Top = 96
    Width = 121
    Height = 23
    Hint = 'Arquivo'
    TabOrder = 6
  end
  object edtModulos: TEdit
    Left = 301
    Top = 96
    Width = 121
    Height = 23
    Hint = 'M'#243'dulos'
    TabOrder = 7
  end
  object edtKWp: TEdit
    Left = 13
    Top = 152
    Width = 121
    Height = 23
    Hint = 'KWp'
    TabOrder = 8
  end
  object edtArea: TEdit
    Left = 156
    Top = 152
    Width = 121
    Height = 23
    Hint = 'Area'
    TabOrder = 9
  end
  object OpenDialog: TOpenDialog
    DefaultExt = '*.xlsx;*.xls'
    Filter = 
      'Planilhas (*.xlsx, *.xls)|*.xlsx;*.xls|Todos os arquivos (*.*)|*' +
      '.*'
    Left = 245
    Top = 65529
  end
end
