""" first big data test: data importing and manipulation"""
__author__ = "Franco Sacchetti"
__copyrights__ = "Copyright 2019, Marelli PWT"
__credits__ = ["Franco Sacchetti"]

__license__ = "MIT"
__maintainer__ = "Franco Sacchetti"
__email__ = "franco.sacchetti@magnetimarelli.com"
__status__ = "Development"

# from time import sleep
import numpy as np  # pragma: no cover
import sys  # pragma: no cover
import pandas as pd  # pragma: no cover
import matplotlib.pyplot as plt  # pragma: no cover


def header_filler(df):
    try:
        # print('DEBUG-->header_filler columns :\n', df.columns)
        df.columns = ['_'.join(c.split()) for c in df.columns]
    except IOError:
        print('DEBUG-->header_filler, exception raised: column not a string', df.columns,
              'sys exception: ', sys.exc_info())
        pass
    return df


def clean_unnamed(df):
    list_to_drop = []
    try:
        for i in range(len(df.columns)):
            if isinstance(df.columns[i], str):
                if 'Unnamed' in df.columns[i] or 'index' in df.columns[i]:
                    list_to_drop.append(str(df.columns[i]))
        # df.drop(list_to_drop, axis=1)
        # print('data to drop: ', list_to_drop)
    except IOError:
        print('DEBUG-->main: exception raised: ', sys.exc_info())
    return list_to_drop


def save_new_excel_data(df, req_file_name, sheet):
    """DataFrame.to_excel(self, excel_writer, sheet_name='Sheet1', na_rep='', float_format=None, columns=None,
    header=True, index=True, index_label=None, startrow=0, startcol=0, engine=None, merge_cells=True, encoding=None,
    inf_rep='inf', verbose=True, freeze_panes=None)[source]"""
    try:
        # select rows for a specific column and save a excel file
        dtc_table_ext = ['SW_DTC', 'Diagnosis_IDENTIFIER', 'Symptom', 'SW_Module', 'ISO_Pcode',
                           'Cust_Pcode', 'ScanT_Pcode', 'Description', 'Lamp_Manager', 'EPC_Lamp',
                           'SnapShot', 'MIL_FUEL_CONF', 'Diagnosis_Enabled', 'Diagnosis_presence',
                           'Severity', 'Priority', 'Diag_Call_task', 'Diag_Validation', 'Unit',
                           'Diag_DeValidation', 'DTC_available', 'EPC', 'MIL_FuelConf_bit1',
                           'MIL_FuelConf_bit0', 'Lamp_Manager_bit2', 'Lamp_Manager_bit1', 'Lamp_Manager_bit0',
                           'AUTOyyy', 'Prio_bit3', 'Prio_bit2', 'Prio_bit1', 'Prio_bit0',
                           'Snapshot_bit2', 'Snapshot_bit1', 'Snapshot_bit0', 'empty', 'ETC_highbit', 'ETC_lowbit']
        # Save df_all_cols extracted to a new excel file
        file_to_save = sheet+'_'+req_file_name
        with pd.ExcelWriter(file_to_save) as writer:  # for writing more than 1 sheet
            df.to_excel(writer, sheet_name=sheet, index=False)
            # df.to_excel(writer, sheet_name=sheet, columns=dtc_table_ext, index=False)
    except PermissionError:
        print('DEBUG-->save_new_excel_data: exception raised: ', sys.exc_info())


def plot_dtc_stats(df):
    # use a plot style
    # seaborn-dark-palette seaborn-dark seaborn-colorblind dark_background seaborn-white fivethirtyeight ggplot
    try:
        plt.style.use('seaborn-whitegrid')  # plots style: seaborn-bright seaborn-whitegrid

        # Company colors
        marelli_text = '#002855'
        marelli_symb = '#009CDE'
        # create Figure_1
        fig = plt.figure(figsize=(10, 9), frameon=False)  # hspace=0.30)
        fig.suptitle('DTC Table stats')
        # fig.tight_layout()
        # Divide the figure into a 2x2 grid
        ax1 = fig.add_subplot(221)
        ax2 = fig.add_subplot(222)
        ax3 = fig.add_subplot(223)
        ax4 = fig.add_subplot(224)

        # subplots parameters
        ax3.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        ax4.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        nich = 0.1  # out of pie
        ich = 0.05  # out of pie
        ax1.set_xlabel('number of', fontsize=8, fontweight='black', color='#323F4B')

        # ['Status', 'Severity', 'Enterer Name'].plot(hatch='/', kind='barh') # marker='o',
        df['Priority'].value_counts().sort_values().plot.barh(alpha=0.8, subplots=True, fontsize=8,
                                                              color=marelli_symb,
                                                              ax=ax1)
        df['Symptom'].value_counts()[:10].plot.bar(alpha=0.8,
                                                   color=marelli_text,
                                                   subplots=True,
                                                   fontsize=8,
                                                   ax=ax2)
        explode3 = (nich, ich, ich, ich)  # , ich)  # , ich, nich, nich, nich, nich)  # , 0, 0, 0, 0)
        # print(df['Priority'].value_counts())
        df['Priority'].value_counts()[:4].plot.pie(explode=explode3,
                                                   autopct='%1.1f%%',
                                                   shadow=True,
                                                   startangle=75,
                                                   label='Priority',
                                                   subplots=False,
                                                   ax=ax3)
        explode6 = (nich, ich, ich, ich)  # ich, ich, ich, ich, nich, nich, nich, nich)  # , 0, 0, 0, 0)
        df['Symptom'].value_counts()[:4].plot.pie(explode=explode6,
                                                  autopct='%1.1f%%',
                                                  shadow=True,
                                                  startangle=45,
                                                  label='Symptom \n top4',
                                                  subplots=False,
                                                  ax=ax4)

        fig1 = plt.figure(figsize=(10, 9))
        ax5 = fig1.add_subplot(221)
        ax6 = fig1.add_subplot(222)
        ax7 = fig1.add_subplot(223)
        ax8 = fig1.add_subplot(224)

        # Equal aspect ratio ensures that pie is drawn as a circle.
        ax5.axis('equal')
        explode5 = (nich, ich)  # , ich, ich, ich, ich, ich, ich, nich, nich)  # , nich, nich)  # , 0, 0, 0, 0)
        # df.groupby(['SW_Module'])['Diagnosis_Enabled'].value_counts()[:10].sort_values().plot.pie(
        # print('DEBUG: Diagnosis_Enabled : \n', df['Diagnosis_Enabled'])
        df['Diagnosis_Enabled'].value_counts().plot.pie(explode=explode5,
                                                        autopct='%1.1f%%',
                                                        shadow=True,
                                                        startangle=45,
                                                        label='Diagnosis_Enabled',
                                                        subplots=False,
                                                        ax=ax5)
        # print('Diag_Call_task : ', df['Diag_Call_task'])
        df['Diag_Call_task'].value_counts().sort_values().plot.barh(alpha=0.8,
                                                                    subplots=True,
                                                                    fontsize=8,
                                                                    label='Diag_Call_task',
                                                                    color=marelli_symb,
                                                                    ax=ax6)
        df['Diag_Validation'].value_counts()[:10].sort_values().plot.barh(alpha=0.8,
                                                                          subplots=True,
                                                                          fontsize=8,
                                                                          label='Diag_Validation top10',
                                                                          color=marelli_symb,
                                                                          ax=ax7)

        df['Diag_DeValidation'].value_counts()[:10].sort_values().plot.barh(alpha=0.8,
                                                                            subplots=True,
                                                                            fontsize=8,
                                                                            label='Diag_DeValidation top10',
                                                                            color=marelli_symb,
                                                                            ax=ax8)

        fig2 = plt.figure(figsize=(10, 7))
        ax9 = fig2.add_subplot(221)
        ax10 = fig2.add_subplot(222)
        ax11 = fig2.add_subplot(223)
        ax12 = fig2.add_subplot(224)

        df['Lamp_Manager'].value_counts().sort_values().plot.barh(alpha=0.8,
                                                                  subplots=True,
                                                                  fontsize=8,
                                                                  label='Lamp_Manager',
                                                                  color=marelli_symb,
                                                                  ax=ax9)
        df['EPC_Lamp'].value_counts().sort_values().plot.barh(alpha=0.8,
                                                              subplots=True,
                                                              fontsize=8,
                                                              label='EPC_Lamp',
                                                              color=marelli_symb,
                                                              ax=ax10)
        df['SnapShot'].value_counts().sort_values().plot.barh(alpha=0.8,
                                                              subplots=True,
                                                              fontsize=8,
                                                              label='SnapShot',
                                                              color=marelli_symb,
                                                              ax=ax11)
        df['MIL_FUEL_CONF'].value_counts().sort_values().plot.barh(alpha=0.8,
                                                                   subplots=True,
                                                                   fontsize=8,
                                                                   label='MIL_FUEL_CONF',
                                                                   color=marelli_symb,
                                                                   ax=ax12)

        # print(df['SDTC_CAL_'].value_counts())
        fig3 = plt.figure(figsize=(10, 7))
        ax13 = fig3.add_subplot(221)
        ax14 = fig3.add_subplot(222)
        ax15 = fig3.add_subplot(223)
        ax16 = fig3.add_subplot(224)
        df.plot.scatter(['HDTC_CAL_'], ['SDTC_CAL_'], alpha=0.8, subplots=True, fontsize=8, label='Scatter Plot',
                        ax=ax13)
        df['DTC_SYM_CAL_'].plot(alpha=0.8, subplots=True, fontsize=8, label='DTC_SYM_CAL_', ax=ax14)
        # df.plot.line(0, ['C_'], alpha=0.8, subplots=True, fontsize=8, label='C_', color=marelli_symb, ax=ax15)
        # df['DIAGCALFLAGS2_bitmask'].plot(alpha=0.8, subplots=True, fontsize=8, label='DIAGCALFLAGS2_bitmask', ax=ax16)

        # Adjust the margins
        plt.subplots_adjust(top=0.92, bottom=0.08, left=0.10, right=0.95, hspace=0.25, wspace=0.35)
        plt.show()
    except ValueError:
        print('DEBUG-->plot_dtc_stats, exception raised: ', sys.exc_info())


def plot_rec_stats(df):
    # use a plot style
    # seaborn-dark-palette seaborn-dark seaborn-colorblind dark_background seaborn-white fivethirtyeight ggplot
    try:
        plt.style.use('seaborn-whitegrid')  # plots style: seaborn-bright seaborn-whitegrid
        # Company colors
        marelli_text = '#002855'
        marelli_symb = '#009CDE'
        # create Figure_1
        fig = plt.figure(figsize=(10, 9), frameon=False)  # hspace=0.30)
        fig.suptitle('REC Table stats top 10')
        # fig.tight_layout()
        # Divide the figure into a 2x1 grid
        ax1 = fig.add_subplot(211)
        # ax2 = fig.add_subplot(212)
        ax3 = fig.add_subplot(223)
        # ax4 = fig.add_subplot(224)

        # subplots parameters
        ax3.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        # ax4.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        nich = 0.1  # out of pie
        ich = 0.05  # out of pie
        ax1.set_xlabel('number of', fontsize=8, fontweight='black', color='#323F4B')

        # ['SW_DTC', 'Diagnosis_IDENTIFIER', 'Development_Tool'].plot(hatch='/', kind='barh') # marker='o',
        df['Development_Tool'].value_counts()[:10].sort_values().plot.barh(alpha=0.9, subplots=True, fontsize=8,
                                                                           color=marelli_symb,
                                                                           ax=ax1)
        explode3 = (nich, ich, ich, ich, ich, ich, nich, nich, nich, nich)  # , 0, 0, 0, 0)
        # print(df['Development_Tool'].value_counts())
        df['Development_Tool'].value_counts()[:10].plot.pie(explode=explode3,
                                                            autopct='%1.1f%%',
                                                            shadow=True,
                                                            startangle=45,
                                                            label='Development_Tool\ntop 10',
                                                            subplots=False,
                                                            ax=ax3)
        # Adjust the margins
        plt.subplots_adjust(top=0.92, bottom=0.08, left=0.15, right=0.95, hspace=0.25, wspace=0.35)
        plt.show()
    except ValueError:
        print('DEBUG-->plot_rec_stats, exception raised: ', sys.exc_info())


def open_req_diag_dtc(file_name, sheet):

    # pandas.read_excel(io, sheet_name=0, header=0, names=None, index_col=None, parse_cols=None, usecols=None,
    # squeeze=False, dtype=None, engine=None, converters=None, true_values=None, false_values=None, skiprows=None,
    # nrows=None, na_values=None, keep_default_na=True, verbose=False, parse_dates=False, date_parser=None,
    # thousands=None, comment=None, skip_footer=0, skipfooter=0, convert_float=True, mangle_dupe_cols=True, **kwds)
    df_all_cols = None
    op_status = ''
    try:
        data = pd.read_excel(file_name, sheet_name=1, header=4)  # [0, 1, "Sheet5"]) index_col=2
        data_2 = pd.read_excel(file_name, sheet_name=1, header=1, skiprows=[2, 3, 4])
        # clean header
        data = header_filler(data)
        data_2 = header_filler(data_2)
        # clean no data columns
        drop_list_columns_data_1 = clean_unnamed(data)
        drop_list_columns_data_2 = clean_unnamed(data_2)
        # drop columns using list of column names in place with same DataFrame
        data.drop(drop_list_columns_data_1, axis=1, inplace=True)
        data_2.drop(drop_list_columns_data_2, axis=1, inplace=True)

        # concatenate all columns of data and data_2
        df_all_cols = pd.concat([data, data_2], axis=1)

        # clear database from non useful rows and columns
        df_all_cols['Diagnosis_IDENTIFIER'].replace('', np.nan, inplace=True)
        df_all_cols.dropna(subset=['Diagnosis_IDENTIFIER'], inplace=True)
        df_all_cols.drop(df_all_cols.index[df_all_cols['Diagnosis_IDENTIFIER'] == 'SPARE'], inplace=True)
        df_all_cols.drop('<==_Insert_DTC_or_Inca_Error_code_to_get_information_about', axis=1, inplace=True)
        df_all_cols.drop('DIAGCALFLAGS_bitmask', axis=1, inplace=True)
        df_all_cols.drop('TOOLCALFLAGS_bitmask', axis=1, inplace=True)
        df_all_cols.drop('DIAGCALFLAGS2_bitmask', axis=1, inplace=True)

        # print info
        print(df_all_cols.head())
        # print(df_all_cols.dtypes)
        print(df_all_cols.describe())
        print(df_all_cols.info())
        # TODO: add data types to dataframe import like df['Lamp_Manager_bit0'].astype('bool')
        # extract a categories index
        # index = list(df_all_cols)
        # print(index)
        # save a excel file with select rows and columns
        save_new_excel_data(df_all_cols, file_name, sheet)
        op_status = 'OK'
    except IOError:
        print('DEBUG-->open_req_diag_dtc, exception raised: ', sys.exc_info())
        op_status = 'Operation KO '+str(sys.exc_info())
    return df_all_cols, op_status


def open_req_diag_rec(file_name, sheet):
    # pandas.read_excel(io, sheet_name=0, header=0, names=None, index_col=None, parse_cols=None, usecols=None,
    # squeeze=False, dtype=None, engine=None, converters=None, true_values=None, false_values=None, skiprows=None,
    # nrows=None, na_values=None, keep_default_na=True, verbose=False, parse_dates=False, date_parser=None,
    # thousands=None, comment=None, skip_footer=0, skipfooter=0, convert_float=True, mangle_dupe_cols=True, **kwds)
    df_all_cols = None
    op_status = ''
    try:
        data_1 = pd.read_excel(file_name, sheet_name='RECMGM', header=4, usecols='B:D')  # [0, 1, "Sheet5"]) index_col=2
        data_2 = pd.read_excel(file_name, sheet_name='RECMGM', header=1, skiprows=[2, 3, 4], usecols='E:BP')
        # clean header
        data_1 = header_filler(data_1)
        data_2 = header_filler(data_2)
        # list coclean no data columns
        drop_list_columns_data_1 = clean_unnamed(data_1)
        drop_list_columns_data_2 = clean_unnamed(data_2)
        # drop columns using list of column names in place with same DataFrame
        data_1.drop(drop_list_columns_data_1, axis=1, inplace=True)
        data_2.drop(drop_list_columns_data_2, axis=1, inplace=True)

        # concatenate all columns of data and data_2
        df_all_cols = pd.concat([data_1, data_2], axis=1)

        # clear database from non useful rows and columns
        df_all_cols['Diagnosis_IDENTIFIER'].replace('', np.nan, inplace=True)
        df_all_cols.dropna(subset=['Diagnosis_IDENTIFIER'], inplace=True)
        df_all_cols.drop(df_all_cols.index[df_all_cols['Diagnosis_IDENTIFIER'] == 'SPARE'], inplace=True)
        # df_all_cols.drop('<==_Insert_DTC_or_Inca_Error_code_to_get_information_about', axis=1, inplace=True)
        # df_all_cols.drop('DIAGCALFLAGS_bitmask', axis=1, inplace=True)
        # df_all_cols.drop('TOOLCALFLAGS_bitmask', axis=1, inplace=True)
        # df_all_cols.drop('DIAGCALFLAGS2_bitmask', axis=1, inplace=True)

        # print info
        print(df_all_cols.head())
        # print(df_all_cols.dtypes)
        print(df_all_cols.describe())
        print(df_all_cols.info())
        # TODO: add data types to dataframe import like df['Lamp_Manager_bit0'].astype('bool')
        # extract a categories index
        # index = list(df_all_cols)
        # print(index)
        # save a excel file with select rows and columns
        save_new_excel_data(df_all_cols, file_name, sheet)
        op_status = 'OK'
    except IOError:
        print('DEBUG-->open_req_diag_rec, exception raised: ', sys.exc_info())
        op_status = 'Operation KO '+str(sys.exc_info())
    return df_all_cols, op_status


def df_diff(old_frame, new_frame, file1, file2):
    diff = None
    op_status = ''
    try:
        if not old_frame.empty:
            df_bool = (old_frame != new_frame).stack()
            diff = pd.concat([old_frame.stack()[df_bool], new_frame.stack()[df_bool]], axis=1)
            diff.index.names = ['DTC #', 'Field']
            # diff.rename(columns={diff.columns[-2]: "your value"}, inplace=True)
            # print(diff.head())
            diff.columns = [file1, file2]
            op_status = 'Operation OK'
        else:
            print('\nDEBUG-->df_diff, not existent dataframe\n')
            op_status = 'Operation KO'
    except IOError:
        print('\nDEBUG-->df_diff, not existent dataframe\n')
        print('DEBUG-->df_diff, exception raised: ', sys.exc_info())
        op_status = 'Operation KO'
    return diff, op_status


def write_excel_df_diff(f_name_1, f_name_2):
    df1, df2, df3, df4 = None, None, None, None
    difference_dtc, difference_rec = None, None
    try:
        # open dtc table and find difference
        df1, op_status = open_req_diag_dtc(f_name_1, 'DTC_Extract')
        df2, op_status = open_req_diag_dtc(f_name_2, 'DTC_Extract')
        if 'KO' not in op_status:
            difference_dtc, op_status = df_diff(df1, df2, f_name_1, f_name_2)
        # open rec table and find difference
        df3, op_status = open_req_diag_rec(f_name_1, 'REC_Extract')
        df4, op_status = open_req_diag_rec(f_name_2, 'REC_Extract')
        if 'KO' not in op_status:
            difference_rec, op_status = df_diff(df3, df4, f_name_1, f_name_2)
        # save differences to excel file 'difference.xlsx'
        if 'KO' not in op_status:
            with pd.ExcelWriter('difference.xlsx') as writer:  # for writing more than 1 sheet
                if difference_dtc.index.size == 0:
                    difference_dtc = difference_dtc.append({f_name_1: 'No modifications',
                                                            f_name_2: 'No modifications'}, ignore_index=True)
                if difference_rec.index.size == 0:
                    difference_rec = difference_rec.append({f_name_1: 'No modifications',
                                                            f_name_2: 'No modifications'}, ignore_index=True)
                difference_dtc.to_excel(writer, sheet_name='Diff_Req_DTC', index=True)
                difference_rec.to_excel(writer, sheet_name='Diff_Req_REC', index=True)
            op_status = 'Operation OK'
        else:
            op_status = 'Operation KO'
    except IOError:
        print('DEBUG-->main: exception raised: ', sys.exc_info())
        op_status = 'Operation KO'+str(sys.exc_info())
    return [df1, df2, df3, df4, difference_dtc, difference_rec, op_status]


# ------------------------------------------------
#   TEST MAIN
# ------------------------------------------------
#
# Run the program
if __name__ == '__main__':
    file_name_1 = 'Req_SYSDIAG_P230_EMEA_20062019.xlsx'
    file_name_2 = 'Req_SYSDIAG_P230_NAFTA_20062019.xlsx'
    [df_1, df_2, df_3, df_4, diff_dtc, diff_rec, operation_status] = \
        write_excel_df_diff(file_name_1, file_name_2)
    # print to console the differences
    print('Operation status: ', operation_status)
    print(diff_dtc)
    print(diff_rec)
    plot_dtc_stats(df_1)
    plot_rec_stats(df_3)
