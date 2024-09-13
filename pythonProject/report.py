import pandas as pd
import numpy as np
import xlsxwriter

def report_creator():
    def mt103():
        try:
            original_103_df = pd.read_excel('uploads/mt103.xls', parse_dates=True).dropna(how='all')
            dropped_cols=['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 8', 'Unnamed: 11']
            rename={'Unnamed: 1':'SENDER', 'Unnamed: 2':'RECEIVER', 'Unnamed: 5':'DATE',
                    'Unnamed: 6':'CURRENCY', 'Unnamed: 7':'AMOUNT', 'Unnamed: 9':'ACCOUNT', 'Unnamed: 10':'REASON', 'Unnamed: 12':'BENEFICIARY'}
            mt103_df = pd.read_excel('uploads/mt103.xls', parse_dates=True).drop(labels=dropped_cols, axis=1).rename(columns=rename).dropna(how='all')
            mt103_df = mt103_df[['SENDER','RECEIVER','DATE','CURRENCY','AMOUNT','BENEFICIARY','REASON','ACCOUNT']]
            mask1 = (mt103_df.REASON.astype('str').str.lower().str.contains('loan'))|(mt103_df['ACCOUNT'].astype('str').str.upper().str.startswith('AA'))
            mt103_df.loc[mask1, 'REASON'] = 'LOAN REPAYMENT'
            mask2 = (mt103_df['SENDER'] == 'CBKEKENXXXX')
            mt103_df.loc[(mt103_df['SENDER'] == 'CBKEKENXXXX'), 'REASON'] = original_103_df['Unnamed: 11']
            mt103_df['REASON']=mt103_df.REASON.astype('str').str.upper().str.replace(r'/REC/|/ROC/|/RFB/|/', '', regex=True)
            return mt103_df
        except FileNotFoundError as e:
            print(f'mt103 error {e}')
            pass

    def mt202():
        try:
            mt202_df = pd.read_excel(r'uploads/mt202.xls', parse_dates=True).dropna(how='any')
            condition = [mt202_df['Unnamed: 7'] == 'USD',mt202_df['Unnamed: 7'] == 'EUR',mt202_df['Unnamed: 7'] == 'GBP',mt202_df['Unnamed: 7'] == 'KES']
            choices = ['1006940478', '1008458398', '1006940467', '1007102778']
            conditions1 = [mt202_df['Unnamed: 4'].astype('str').str.upper().str.startswith('FX'),
                           mt202_df['Unnamed: 4'].astype('str').str.upper().str.contains(r'CUSTODY'),
                           mt202_df['Unnamed: 4'].astype('str').str.upper().str.contains(r'BORROW'),
                           mt202_df['Unnamed: 4'].astype('str').str.upper().str.startswith(r'RTN')
                          ]
            choices1 = ['FX DEAL', 'PLACEMENT', 'BORROWING', 'RETURN OF FUNDS']
            mt202_df = (mt202_df.assign(SENDER=mt202_df['Unnamed: 1'],RECEIVER=mt202_df['Unnamed: 2'],DATE=mt202_df['Unnamed: 6'],CURRENCY=mt202_df['Unnamed: 7'],
                                       AMOUNT=mt202_df['Unnamed: 8'].astype('float'),BENEFICIARY='FAULU MICROFINANCE BANK',
                                       REASON=np.select(conditions1,choices1,default='BANK TRANSFERS'),ACCOUNT=np.select(condition, choices,default='KES'))
                        .drop(labels=[col for col in mt202_df.columns if col.startswith('Unnamed')],axis=1))
            mt202_df = mt202_df[['SENDER','RECEIVER','DATE','CURRENCY','AMOUNT','BENEFICIARY','REASON','ACCOUNT']]
            return mt202_df
        except FileNotFoundError as e:
            print(f'mt202 error {e}')
            pass

    def mt102():
        try:
            mt102_df = pd.read_excel(r'uploads/mt102.xls', parse_dates=True).dropna(how='any')
            mt102_df = (mt102_df.assign(SENDER=mt102_df['Unnamed: 1'],RECEIVER=mt102_df['Unnamed: 2'],DATE=mt102_df['Unnamed: 6'],CURRENCY=mt102_df['Unnamed: 7'],
                                      AMOUNT=mt102_df['Unnamed: 8'].astype('float'),BENEFICIARY='FAULU MICROFINANCE BANK',REASON='MINISTRIES',ACCOUNT='1007102778')
                        .drop(labels=[col for col in mt102_df.columns if col.startswith('Unnamed')], axis=1))
            mt102_df = mt102_df[['SENDER','RECEIVER','DATE','CURRENCY','AMOUNT','BENEFICIARY','REASON','ACCOUNT']]
            return mt102_df
        except FileNotFoundError as e:
            print(f'mt102 error {e}')
            pass

    def mt910():
        try:
            mt910_df = pd.read_excel(r'uploads/mt910.xls', parse_dates=True).dropna(how='any')
            condition = [mt910_df['Unnamed: 4'].astype('str').str.upper().str.contains('IPSL')]
            choices = ['PESALINK SETTLEMENT']
            mt910_df = (mt910_df.assign(SENDER=mt910_df['Unnamed: 1'],RECEIVER=mt910_df['Unnamed: 2'],DATE=mt910_df['Unnamed: 6'],CURRENCY=mt910_df['Unnamed: 7'],
                                       AMOUNT=mt910_df['Unnamed: 8'].astype('float'),BENEFICIARY='FAULU MICROFINANCE BANK',
                                       REASON=np.select(condition, choices, default='CBK SETTLEMENTS'), ACCOUNT='1007102778')
                        .drop(labels=[col for col in mt910_df.columns if col.startswith('Unnamed')], axis=1))
            mt910_df = mt910_df[['SENDER','RECEIVER','DATE','CURRENCY','AMOUNT','BENEFICIARY','REASON','ACCOUNT']]
            return mt910_df
        except FileNotFoundError as e:
            print(f'mt910 error {e}')
            pass

    def mt103out():
        try:
            mt103out_df= pd.read_excel(r'uploads/mt103out.xls', parse_dates=True).dropna(how='any')
            mt103out_df= (mt103out_df.assign(SENDER = mt103out_df['Unnamed: 1'],RECEIVER = mt103out_df['Unnamed: 2'],DATE = mt103out_df['Unnamed: 5'],
                                            CURRENCY = mt103out_df['Unnamed: 6'],AMOUNT = mt103out_df['Unnamed: 7'].astype('float'),
                                            BENEFICIARY = mt103out_df['Unnamed: 12'],REASON = mt103out_df['Unnamed: 10'],
                                            ACCOUNT = mt103out_df['Unnamed: 8'].astype('object'))
                          .drop(labels=[col for col in mt103out_df.columns if col.startswith('Unnamed')],axis=1))
            mask3 = mt103out_df['REASON'].str.lower().str.contains('loan')
            mt103out_df.loc[mask3, 'REASON'] = 'LOAN REPAYMENT'
            mask4 = mt103out_df['ACCOUNT']=='KES1050000010001'
            mt103out_df.loc[mask4, 'REASON'] = 'SOLVE DISBURSEMENT'
            mt103out_df = mt103out_df[['SENDER','RECEIVER','DATE','CURRENCY','AMOUNT','BENEFICIARY','REASON','ACCOUNT']]
            return mt103out_df
        except FileNotFoundError as e:
            print(f'mt103out error {e}')
            pass

    def mt202out():
        try:
            mt202out = pd.read_excel(r'uploads/mt202out.xls', parse_dates=True).dropna(how='any')
            condition = [mt202out['Unnamed: 7'] == 'USD',mt202out['Unnamed: 7'] == 'EUR',mt202out['Unnamed: 7'] == 'GBP',mt202out['Unnamed: 7'] == 'KES']
            choices = ['1006940478', '1008458398', '1006940467', '1007102778']
            conditions1 = [mt202out['Unnamed: 3'].str.upper().str.startswith('FX'),
                           mt202out['Unnamed: 3'].str.upper().str.startswith('MM'),
                           mt202out['Unnamed: 3'].str.upper().str.startswith('RTN')]
            choices1 = ['FX DEAL', 'PLACEMENT/BORROWING', 'RETURNED FUNDS']
            mt202out = mt202out.assign(SENDER = mt202out['Unnamed: 1'],RECEIVER = mt202out['Unnamed: 2'],DATE = mt202out['Unnamed: 6'],
                                       CURRENCY = mt202out['Unnamed: 7'],AMOUNT = mt202out['Unnamed: 8'].astype('float'),BENEFICIARY = 'FAULU MICROFINANCE BANK',
                                       REASON = np.select(conditions1,choices1,default='BANK TRANSFERS'),ACCOUNT = np.select(condition, choices, default='KES')
                                      ).drop(labels=[col for col in mt202out.columns if col.startswith('Unnamed')],axis=1)
            mt202out = mt202out[['SENDER','RECEIVER','DATE','CURRENCY','AMOUNT','BENEFICIARY','REASON','ACCOUNT']]
            return mt202out
        except FileNotFoundError as e:
            print(f'mt202out error {e}')
            pass

    def mt900out():
        try:
            mt900out_df = pd.read_excel(r'uploads/mt900out.xls', parse_dates=True).dropna(how='any')
            condition1 = [mt900out_df['Unnamed: 7'] == 'USD',mt900out_df['Unnamed: 7'] == 'EUR',mt900out_df['Unnamed: 7'] == 'GBP',
                          mt900out_df['Unnamed: 7'] == 'KES']
            choices1 = ['1006940478', '1008458398', '1006940467', '1007102778']
            condition=[mt900out_df['Unnamed: 4'].str.upper().str.startswith('KES'),mt900out_df['Unnamed: 4'].str.upper().str.startswith('BILL')]
            choices=['PESALINK SETTLEMENT', 'CBK CHARGES']
            mt900out_df = (mt900out_df.assign(SENDER = mt900out_df['Unnamed: 2'],RECEIVER = mt900out_df['Unnamed: 1'],DATE = mt900out_df['Unnamed: 6'],
                                             CURRENCY = mt900out_df['Unnamed: 7'], AMOUNT = mt900out_df['Unnamed: 8'].astype('float'),BENEFICIARY = 'CENTRAL BANK',
                                             REASON = np.select(condition,choices,default='CBK PAYMENTS'),
                                             ACCOUNT = np.select(condition1, choices1, default='KES'))
                           .drop(labels=[col for col in mt900out_df.columns if col.startswith('Unnamed')],axis=1))
            mt900out_df = mt900out_df[['SENDER','RECEIVER','DATE','CURRENCY','AMOUNT','BENEFICIARY','REASON','ACCOUNT']]
            return mt900out_df
        except FileNotFoundError as e:
            print(f'mt900out error {e}')
            pass

    merged_incomings_df = None
    merged_outgoings_df = None
    try:
        merged_incomings_df = pd.concat([mt103(), mt202(), mt102(), mt910()]).sort_values(by='AMOUNT').reset_index(drop=True)
        merged_outgoings_df = pd.concat([mt103out(), mt202out(), mt900out()]).sort_values(by='AMOUNT').reset_index(drop=True)
    except ValueError as e:
        print(f'merged dataframe empty: {e}')
        pass

    kes_incoming_df= None
    kes_outgoing_df= None
    kes_ind = None
    eur_incoming_df= None
    eur_outgoing_df= None
    eur_ind = None
    usd_incoming_df= None
    usd_outgoing_df= None
    usd_ind = None
    gbp_incoming_df= None
    gbp_outgoing_df= None
    gbp_ind = None
    loan_payment_incoming = None
    loan_payment_outgoing = None
    loan_ind = None
    non_loan_payment_incoming = None
    non_loan_payment_outgoing = None
    non_loan_ind = None
    locking_df = None

    if merged_incomings_df.empty:
        print('merged_incomings_df is empty')
        pass
    else:
        kes_incoming_df = merged_incomings_df[merged_incomings_df['CURRENCY'] == 'KES']
        usd_incoming_df = merged_incomings_df[merged_incomings_df['CURRENCY'] == 'USD']
        eur_incoming_df = merged_incomings_df[merged_incomings_df['CURRENCY'] == 'EUR']
        gbp_incoming_df = merged_incomings_df[merged_incomings_df['CURRENCY'] == 'GBP']
        loan_payment_incoming = kes_incoming_df[kes_incoming_df['REASON']=='LOAN REPAYMENT']
        non_loan_payment_incoming = kes_incoming_df[kes_incoming_df['REASON']!='LOAN REPAYMENT']
        kes_ind = len(kes_incoming_df['AMOUNT'])
        usd_ind = len(usd_incoming_df['AMOUNT'])
        eur_ind = len(eur_incoming_df['AMOUNT'])
        gbp_ind = len(gbp_incoming_df['AMOUNT'])
        non_loan_ind = len(non_loan_payment_incoming['AMOUNT'])
        loan_ind = len(loan_payment_incoming['AMOUNT'])
        print('unable to find match')

    if merged_outgoings_df.empty:
        print('merged_outgoing_df is empty')
        pass
    else:
        kes_outgoing_df = merged_outgoings_df[merged_outgoings_df['CURRENCY'] == 'KES']
        usd_outgoing_df = merged_outgoings_df[merged_outgoings_df['CURRENCY'] == 'USD']
        eur_outgoing_df = merged_outgoings_df[merged_outgoings_df['CURRENCY'] == 'EUR']
        gbp_outgoing_df = merged_outgoings_df[merged_outgoings_df['CURRENCY'] == 'GBP']
        loan_payment_outgoing = kes_outgoing_df[kes_outgoing_df['REASON']=='LOAN REPAYMENT']
        non_loan_payment_outgoing = kes_outgoing_df[kes_outgoing_df['REASON']!='LOAN REPAYMENT']
        mask6=(kes_incoming_df['AMOUNT']>=1000000)
        mask7=(kes_incoming_df['BENEFICIARY'].astype('str').str.lower().str.contains('faulu|dry|nwrealite|mutual|nominees'))
        locking_df = kes_incoming_df.loc[~mask7 & mask6]



        with pd.ExcelWriter(r'RTGS REPORT.xlsx', engine='xlsxwriter', mode='w') as writer:
            if kes_outgoing_df.empty & kes_incoming_df.empty:
                pass

            else:
                kes_incoming_df.to_excel(writer, sheet_name='KES REPORT', startrow=1, index=False)
                kes_outgoing_df.to_excel(writer, sheet_name='KES REPORT', startrow=(kes_ind + 6), index=False)

            if usd_incoming_df.empty & usd_outgoing_df.empty:
                pass
            else:
                usd_incoming_df.to_excel(writer, sheet_name='USD REPORT', startrow=1, index=False)
                usd_outgoing_df.to_excel(writer, sheet_name='USD REPORT', startrow=(usd_ind + 6), index=False)

            if eur_incoming_df.empty & eur_outgoing_df.empty:
                pass
            else:
                eur_incoming_df.to_excel(writer, sheet_name='EUR REPORT', startrow=1, index=False)
                eur_outgoing_df.to_excel(writer, sheet_name='EUR REPORT', startrow=(eur_ind + 6), index=False)

            if gbp_incoming_df.empty & gbp_outgoing_df.empty:
                pass
            else:
                gbp_incoming_df.to_excel(writer, sheet_name='GBP REPORT', startrow=1, index=False)
                gbp_outgoing_df.to_excel(writer, sheet_name='GBP REPORT', startrow=(gbp_ind + 6), index=False)

            if loan_payment_incoming.empty & loan_payment_outgoing.empty:
                pass
            else:
                loan_payment_incoming.to_excel(writer, sheet_name='LOAN PAYMENTS', startrow=1, index=False)
                loan_payment_outgoing.to_excel(writer, sheet_name='LOAN PAYMENTS', startrow=(loan_ind + 6), index=False)

            if non_loan_payment_incoming.empty & non_loan_payment_outgoing.empty:
                pass
            else:
                non_loan_payment_incoming.to_excel(writer, sheet_name='NON-LOAN PAYMENTS', startrow=1, index=False)
                non_loan_payment_outgoing.to_excel(writer, sheet_name='NON-LOAN PAYMENTS', startrow=(non_loan_ind + 6),
                                                   index=False)

            locking_df.to_excel(writer, sheet_name='LOCKED FUNDS', startrow=1, index=False)

            workbook = writer.book
            header_style = {'font_name': 'Garamond', 'font_size': 11, 'border': 0, 'align': 'left', 'bold': True}
            styles = {'font_name': 'Garamond', 'font_size': 11, 'border': 0, 'align': 'left'}
            styles1 = {'font_name': 'Garamond', 'font_size': 11, 'border': 0, 'align': 'left', 'num_format': '#,##0.00'}
            styles2 = {'bold': True, 'font_name': 'Garamond', 'font_size': 11, 'border': 0, 'align': 'left',
                       'num_format': '#,##0.00'}
            format_cell = workbook.add_format(styles)
            format_amount = workbook.add_format(styles1)
            format_header = workbook.add_format(header_style)
            format_total = workbook.add_format(styles2)
            try:
                kes_worksheet = writer.sheets['KES REPORT']
                kes_worksheet.set_zoom(90)
                kes_worksheet.set_column('A:D', 18, format_cell)
                kes_worksheet.set_column('E:E', 15, format_amount)
                kes_worksheet.set_column('F:G', 40, format_amount)
                kes_worksheet.set_column('H:H', 15, format_amount)
                kes_worksheet.write(int(0), int(0), 'INFLOWS', format_header)
                kes_worksheet.write_formula(f'E{kes_ind + 4}', f'=SUM(E3:E{kes_ind + 3})', format_total)
                kes_worksheet.write_formula(f'E{(kes_ind + 7) + (len(kes_outgoing_df) + 2)}',
                                            f'=SUM(E{kes_ind + 8}:E{(kes_ind + 7) + (len(kes_outgoing_df))})'
                                            , format_total)
                kes_worksheet.write((kes_ind + 5), 0, 'OUTFLOWS', format_header)
                for col_num, value in enumerate(kes_incoming_df.columns.values):
                    kes_worksheet.write(1, col_num, value, format_header)
                    kes_worksheet.write((kes_ind + 6), col_num, value, format_header)
                kes_worksheet.set_tab_color('purple')
            except KeyError as e:
                print(f'sheet not found: {e}')
                pass

            try:
                eur_worksheet = writer.sheets['EUR REPORT']
                eur_worksheet.set_zoom(90)
                eur_worksheet.set_column('A:D', 18, format_cell)
                eur_worksheet.set_column('E:E', 15, format_amount)
                eur_worksheet.set_column('F:G', 40, format_amount)
                eur_worksheet.set_column('H:H', 15, format_amount)
                eur_worksheet.write(0, 0, 'INFLOWS', format_header)
                eur_worksheet.write_formula(f'E{eur_ind + 4}', f'=SUM(E3:E{eur_ind + 3})', format_total)
                eur_worksheet.write_formula(f'E{(eur_ind + 7) + (len(eur_outgoing_df) + 2)}',
                                            f'=SUM(E{eur_ind + 8}:E{(eur_ind + 7) + (len(eur_outgoing_df))})'
                                            , format_total)
                eur_worksheet.write((eur_ind + 5), 0, 'OUTFLOWS', format_header)
                for col_num, value in enumerate(kes_incoming_df.columns.values):
                    eur_worksheet.write(1, col_num, value, format_header)
                    eur_worksheet.write((eur_ind + 6), col_num, value, format_header)
                eur_worksheet.set_tab_color('green')
            except KeyError as e:
                print(f'sheet not found: {e}')
                pass

            try:
                gbp_worksheet = writer.sheets['GBP REPORT']
                gbp_worksheet.set_zoom(90)
                gbp_worksheet.set_column('A:D', 18, format_cell)
                gbp_worksheet.set_column('E:E', 15, format_amount)
                gbp_worksheet.set_column('F:G', 40, format_amount)
                gbp_worksheet.set_column('H:H', 15, format_amount)
                gbp_worksheet.write(0, 0, 'INFLOWS', format_header)
                gbp_worksheet.write_formula(f'E{gbp_ind + 4}', f'=SUM(E3:E{gbp_ind + 3})', format_total)
                gbp_worksheet.write_formula(f'E{(gbp_ind + 7) + (len(gbp_outgoing_df) + 2)}',
                                            f'=SUM(E{gbp_ind + 8}:E{(gbp_ind + 7) + (len(gbp_outgoing_df))})'
                                            , format_total)
                gbp_worksheet.write((gbp_ind + 5), 0, 'OUTFLOWS', format_header)
                for col_num, value in enumerate(kes_incoming_df.columns.values):
                    gbp_worksheet.write(1, col_num, value, format_header)
                    gbp_worksheet.write((gbp_ind + 6), col_num, value, format_header)
                gbp_worksheet.set_tab_color('blue')
            except KeyError as e:
                print(f'sheet not found: {e}')
                pass

            try:
                usd_worksheet = writer.sheets['USD REPORT']
                usd_worksheet.set_zoom(90)
                usd_worksheet.set_column('A:D', 18, format_cell)
                usd_worksheet.set_column('E:E', 15, format_amount)
                usd_worksheet.set_column('F:G', 40, format_amount)
                usd_worksheet.set_column('H:H', 15, format_amount)
                usd_worksheet.write(0, 0, 'INFLOWS', format_header)
                usd_worksheet.write_formula(f'E{usd_ind + 4}', f'=SUM(E3:E{usd_ind + 3})', format_total)
                usd_worksheet.write_formula(f'E{(usd_ind + 7) + (len(usd_outgoing_df) + 2)}',
                                            f'=SUM(E{usd_ind + 8}:E{(usd_ind + 7) + (len(usd_outgoing_df))})'
                                            , format_total)
                usd_worksheet.write((usd_ind + 5), 0, 'OUTFLOWS', format_header)
                for col_num, value in enumerate(kes_incoming_df.columns.values):
                    usd_worksheet.write(1, col_num, value, format_header)
                    usd_worksheet.write((usd_ind + 6), col_num, value, format_header)
                usd_worksheet.set_tab_color('lime')
            except KeyError as e:
                print(f'sheet not found: {e}')
                pass

            try:
                loan_worksheet = writer.sheets['LOAN PAYMENTS']
                loan_worksheet.set_zoom(90)
                loan_worksheet.set_column('A:D', 18, format_cell)
                loan_worksheet.set_column('E:E', 15, format_amount)
                loan_worksheet.set_column('F:G', 40, format_amount)
                loan_worksheet.set_column('H:H', 15, format_amount)
                loan_worksheet.write(0, 0, 'INFLOWS', format_header)
                loan_worksheet.write_formula(f'E{loan_ind + 4}', f'=SUM(E3:E{loan_ind + 3})', format_total)
                loan_worksheet.write_formula(f'E{(loan_ind + 7) + (len(loan_payment_outgoing) + 2)}',
                                             f'=SUM(E{loan_ind + 8}:E{(loan_ind + 7) + (len(loan_payment_outgoing))})'
                                             , format_total)
                loan_worksheet.write((loan_ind + 5), 0, 'OUTFLOWS', format_header)
                for col_num, value in enumerate(kes_incoming_df.columns.values):
                    loan_worksheet.write(1, col_num, value, format_header)
                    loan_worksheet.write((loan_ind + 6), col_num, value, format_header)
                loan_worksheet.set_tab_color('orange')
            except KeyError as e:
                print(f'sheet not found: {e}')
                pass

            try:
                non_loan_worksheet = writer.sheets['NON-LOAN PAYMENTS']
                non_loan_worksheet.set_zoom(90)
                non_loan_worksheet.set_column('A:D', 18, format_cell)
                non_loan_worksheet.set_column('E:E', 15, format_amount)
                non_loan_worksheet.set_column('F:G', 40, format_amount)
                non_loan_worksheet.set_column('H:H', 15, format_amount)
                non_loan_worksheet.write(0, 0, 'INFLOWS', format_header)
                non_loan_worksheet.write_formula(f'E{non_loan_ind + 4}', f'=SUM(E3:E{non_loan_ind + 3})', format_total)
                non_loan_worksheet.write_formula(f'E{(non_loan_ind + 7) + (len(non_loan_payment_outgoing) + 2)}',
                                                 f'=SUM(E{non_loan_ind + 8}:E{(non_loan_ind + 7) + (len(non_loan_payment_outgoing))})'
                                                 , format_total)
                non_loan_worksheet.write((non_loan_ind + 6), 0, 'OUTFLOWS', format_header)
                for col_num, value in enumerate(kes_incoming_df.columns.values):
                    non_loan_worksheet.write(1, col_num, value, format_header)
                    non_loan_worksheet.write((non_loan_ind + 6), col_num, value, format_header)
                non_loan_worksheet.set_tab_color('cyan')
            except KeyError as e:
                print(f'sheet not found: {e}')
                pass

            try:
                locked_worksheet = writer.sheets['LOCKED FUNDS']
                locked_worksheet.set_zoom(90)
                locked_worksheet.set_column('A:D', 18, format_cell)
                locked_worksheet.set_column('E:E', 15, format_amount)
                locked_worksheet.set_column('F:G', 40, format_amount)
                locked_worksheet.set_column('H:H', 15, format_amount)
                locked_worksheet.write(0, 0, 'INFLOWS', format_header)
                for col_num, value in enumerate(kes_incoming_df.columns.values):
                    locked_worksheet.write(1, col_num, value, format_header)
                locked_worksheet.set_tab_color('brown')
            except KeyError as e:
                print(f'sheet not found: {e}')
                pass
    return 'RTGS REPORT.xlsx'

