from tabula.io import read_pdf
import tabula
import xlsxwriter
import numpy as np
import pandas as pd

class simas_csv_converter:
  def __init__(self, csv_file, kode_bank, no_virtual_account):
    self.csv_file = csv_file
    self.kode_bank = kode_bank
    self.no_virtual_account = no_virtual_account
  
  def convert(self):
    col_name=['Transaction Reference Number', 'Post Date', 'Transaction Date', 'Transaction Type', 'Description', 'Detail', 'Debit Amount', 'Credit Amount', 'Ending Balance']
    used_col=['Transaction Date', 'Transaction Reference Number', 'Description', 'Debit Amount', 'Credit Amount', 'Ending Balance']
    df = pd.read_csv(self.csv_file, index_col=False, sep=",", skiprows=9, engine='python', header=None, skipfooter=3, names=col_name, usecols=used_col, encoding="utf-8", skipinitialspace=True).dropna(how='all').reset_index(drop=True)

    def f(row):
      if row['Credit Amount'] != np.nan:
        val = 'C'
      else:
        val = 'D'
      return val
        
    df['D/C'] = np.where(df['Credit Amount'].isna(), 'D', 'C')
    df['Amount'] = np.where(df['Credit Amount'].isna(), df['Debit Amount'], df['Credit Amount'])
    df['No. Virtual Account'] = self.no_virtual_account
    df['Kode Bank'] = self.kode_bank
    newName = df.rename(columns={
      'Transaction Reference Number': 'Reference',
      'Description':'description',
      'Transaction Date':'Tanggal Transaksi',
      'Ending Balance': 'Amount in Bank'
    })
    used_cols = newName[['Tanggal Transaksi', 'Reference', 'description', 'No. Virtual Account', 'Amount', 'D/C', 'Amount in Bank', 'Kode Bank']].dropna(how='any')

    writer = pd.ExcelWriter('template_simas.xlsx', engine='xlsxwriter')
    used_cols.to_excel(writer, sheet_name='template_simas')
    writer.save()

class mandiri_converter:
  def __init__(self, pdf_file, kode_bank, no_virtual_account):
    self.pdf_file = pdf_file
    self.kode_bank = kode_bank
    self.no_virtual_account = no_virtual_account
  
  def convert(self):
    data_file = tabula.read_pdf(self.pdf_file, lattice=True, pages='all')
    mandiri_tables = pd.concat(data_file)
    mandiri_tables.reset_index(drop=True, inplace=True)
    kode_bank = self.kode_bank
    no_virtual_account = self.no_virtual_account
    new_name_cols = mandiri_tables.rename(columns={
      'Value Date': "Tanggal Transaksi",
      'Description': "description",
      'Reference No.': 'Reference',
      'Debit': 'Debit Amount',
      'Credit': 'Credit Amount',
      'Saldo': 'Amount in Bank'
    })

    new_name_cols['Kode Bank'] = kode_bank
    new_name_cols['No. Virtual Account'] = no_virtual_account
    new_name_cols['D/C'] = np.where(new_name_cols['Credit Amount'] == 0.00, 'D', 'C')
    new_name_cols['Amount'] = np.where(new_name_cols['Credit Amount'] == 0.00, new_name_cols['Debit Amount'], new_name_cols['Credit Amount'])
    used_cols = new_name_cols[['Tanggal Transaksi', 'Reference', 'description', 'No. Virtual Account', 'Amount', 'D/C', 'Amount in Bank', 'Kode Bank']]

    writer = pd.ExcelWriter('template_mandiri.xlsx', engine='xlsxwriter')
    used_cols.to_excel(writer, sheet_name='template_mandiri')
    writer.save()

class bni_converter:
  def __init__(self, no_virtual_account, kode_bank, pdf_file):
    self.no_virtual_account = no_virtual_account
    self.kode_bank = kode_bank
    self.pdf_file = pdf_file
  
  def convert(self):
    data_file = tabula.read_pdf(self.pdf_file, pages='all', pandas_options={'header': None})
    df = pd.concat(data_file)
    bni_tables = df[df[1].str.contains("Post Date") == False]
    

    new_name_cols = bni_tables.iloc[3:].dropna(how='any').reset_index(drop=True).rename(columns={
        1: 'Tanggal Transaksi',
        3:'Reference',
        4:'description',
        5:'Amount',
        6: 'D/C',
        7: 'Amount in Bank'
      })
    new_name_cols['Kode Bank'] = self.kode_bank
    new_name_cols['No. Virtual Account'] = self.no_virtual_account
    used_cols = new_name_cols[['Tanggal Transaksi', 'Reference', 'description', 'No. Virtual Account', 'Amount', 'D/C', 'Amount in Bank', 'Kode Bank']]

    writer = pd.ExcelWriter('template_bni.xlsx', engine='xlsxwriter')
    used_cols.to_excel(writer, sheet_name='template_bni')
    writer.save()

# test = bni_converter( pdf_file='bni.pdf', kode_bank='', no_virtual_account='')
# test.convert()

# test_mandiri = mandiri_converter(pdf_file='mandiri.pdf', kode_bank='', no_virtual_account='')
# test_mandiri.convert()

test_simas = simas_csv_converter(csv_file='simas_rk.csv', kode_bank='', no_virtual_account='')
test_simas.convert()

