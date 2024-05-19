import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from tkinter import Canvas, Scrollbar, Frame
import seaborn as sns
from tkinter import Tk, Label, Button, filedialog, simpledialog, Frame
from natsort import natsorted
from functools import reduce
from tkinter import *
import customtkinter as ctk  
from tkinter import ttk
from sklearn.decomposition import PCA
import numpy as np
import sys, os
import pyuff
import dataset_18 as d18
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

def rename_columns(df, prefix):
    col_dict = {'Untitled_Time*': 'Frequency'}
    for i in range(5):
        col_dict[f'Untitled'] = f'{prefix}_{1}'
        col_dict[f'Untitled.{i}'] = f'{prefix}_{i+1}'
    df.rename(columns=col_dict, inplace=True)

def rename_columns(df, prefix):
    col_dict = {'Untitled_Time*': 'Frequency'}
    for i in range(5):
        col_dict[f'Untitled'] = f'{prefix}_{1}'
        col_dict[f'Untitled.{i}'] = f'{prefix}_{i+1}'
    df.rename(columns=col_dict, inplace=True)

def rename_columns_fea(df, prefix):
    col_dict = {'Frequency (Hz)': 'Frequency'}
    df.drop(f'Phase (deg)', axis=1, inplace=True)

    for i in range(5):
        col_dict[f'Magnitude'] = f'{prefix}_{1}'
        col_dict[f'Magnitude.{i}'] = f'{prefix}_{i+1}'
    df.rename(columns=col_dict, inplace=True)
    
    # Remove the 'Phase.{i}' columns
    for i in range(5):
        if f'Phase (deg).{i}' in df.columns:
            df.drop(f'Phase (deg).{i}', axis=1, inplace=True)

class MainGUI:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1100x580')
        self.root.title("Preprocessing FRF to WCC")
        self.root.iconbitmap('C:/Users/dania/OneDrive/Documents/UM/UM FILES/SEM 1 2324/FYP/PCA-WCC/trueapp/logo.ico')

        ctk.set_default_color_theme('blue')
        self.sidebar = ctk.CTkFrame(root, width=140, corner_radius=0)
        self.sidebar.pack(fill='y', side='left')

        self.main = ctk.CTkFrame(root)
        self.main.pack(fill='both', expand=True)

        self.frf_combiner_button = ctk.CTkButton(self.sidebar, text="Combine FRF", command=self.open_frf_combiner)
        self.frf_combiner_button.pack(fill='x')

        self.pca_button = ctk.CTkButton(self.sidebar, text="PCA Transformation", command=self.open_pca)
        self.pca_button.pack(fill='x')

        self.wcc_button = ctk.CTkButton(self.sidebar, text="WCC Transformation", command=self.open_wcc)
        self.wcc_button.pack(fill='x')
        
        self.unv_button = ctk.CTkButton(self.sidebar, text="UNV Transformation", command=self.open_unv)
        self.unv_button.pack(fill='x')

        self.feafrf_combiner_button = ctk.CTkButton(self.sidebar, text="FEA Combine FRF", command=self.open_feafrf_combiner)
        self.feafrf_combiner_button.pack(fill='x')

        self.frf_combiner_frame = ctk.CTkFrame(self.main, bg_color='light green')
        self.frf_combiner_gui = FRFCombinerGUI(self.frf_combiner_frame)

        self.pca_frame = ctk.CTkFrame(self.main, bg_color='light yellow')
        self.pca_gui = PCAGUI(self.pca_frame)

        self.wcc_frame = ctk.CTkFrame(self.main, bg_color='light blue')
        self.wcc_gui = WCCGUI(self.wcc_frame)

        self.unv_frame = ctk.CTkFrame(self.main, bg_color='light blue')
        self.unv_gui = UNVGUI(self.unv_frame)

        self.feafrf_combiner_frame = ctk.CTkFrame(self.main, bg_color='light green')
        self.feafrf_combiner_gui = FEAFRFCombinerGUI(self.feafrf_combiner_frame)


    def open_frf_combiner(self):
        self.frf_combiner_frame.pack(fill='both', expand=True)
        self.wcc_frame.pack_forget()
        self.pca_frame.pack_forget()
        self.unv_frame.pack_forget()
        self.feafrf_combiner_frame.pack_forget()


    def open_pca(self):
        self.pca_frame.pack(fill='both', expand=True)
        self.frf_combiner_frame.pack_forget()
        self.wcc_frame.pack_forget()
        self.unv_frame.pack_forget()
        self.feafrf_combiner_frame.pack_forget()


    def open_wcc(self):
        self.wcc_frame.pack(fill='both', expand=True)
        self.frf_combiner_frame.pack_forget()
        self.pca_frame.pack_forget()
        self.unv_frame.pack_forget()
        self.feafrf_combiner_frame.pack_forget()


    def open_unv(self):
        self.unv_frame.pack(fill='both', expand=True)
        self.frf_combiner_frame.pack_forget()
        self.pca_frame.pack_forget()
        self.wcc_frame.pack_forget()
        self.feafrf_combiner_frame.pack_forget()

    def open_feafrf_combiner(self):
        self.feafrf_combiner_frame.pack(fill='both', expand=True)
        self.frf_combiner_frame.pack_forget()
        self.wcc_frame.pack_forget()
        self.pca_frame.pack_forget()
        self.unv_frame.pack_forget()
        


### Code for WCC Transformation
class WCCGUI:
    def __init__(self, root):
        self.root = root
        self.file_paths = []
        self.export_dir = ''

        self.header = ctk.CTkLabel(root, text="Write Severity Level", font=("Arial", 14))
        self.header.pack(pady=10)  # Adjust padding as needed

        self.severity_level = StringVar()
        self.severity_level_entry = ctk.CTkEntry(root, textvariable=self.severity_level, justify='center')
        self.severity_level_entry.pack(fill='x', padx=50, pady=1)
        self.severity_level_entry.bind('<Return>', self.save_severity_level)

        self.select_files_button = ctk.CTkButton(root, text="Select Excel Files", command=self.select_files)
        self.select_files_button.pack(fill='both', expand=True, pady=5)

        self.select_export_dir_button = ctk.CTkButton(root, text="Select Destination Folder", command=self.select_export_dir)
        self.select_export_dir_button.pack(fill='both', expand=True, pady=5)

        self.run_button = ctk.CTkButton(root, text="Run", command=self.run_code)
        self.run_button.pack(fill='both', expand=True, pady=5)

    def save_severity_level(self):
        self.severity_level_value = self.severity_level.get()

    def select_files(self):
        self.file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        print(self.file_paths)

    def select_export_dir(self):
        self.export_dir = filedialog.askdirectory()
        print(self.export_dir)
    
    def display_graphs(self, final_df):
        # Create seaborn bar plots for final_df
        fig1, ax1 = plt.subplots()
        sns.barplot(x='severity', y='wcc', hue='point', data=final_df[final_df['mode_shape'] == 1], ax=ax1)
        fig2, ax2 = plt.subplots()
        sns.barplot(x='severity', y='wcc', hue='point', data=final_df[final_df['mode_shape'] == 2], ax=ax2)
        fig3, ax3 = plt.subplots()
        sns.barplot(x='severity', y='wcc', hue='point', data=final_df[final_df['mode_shape'] == 3], ax=ax3)
        # Create a Frame
        frame = Frame(self.root)
        frame.pack()

        # Create a Canvas
        canvas = Canvas(frame, width=800, height=600)
        canvas.pack(side="left")

    # Add a Scrollbar to the Frame
        scrollbar = Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        # Configure the Canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

        # Create an inner Frame
        inner_frame = Frame(canvas)
        canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        # Create FigureCanvasTkAgg objects for the plots and add them to the inner Frame
        canvas1 = FigureCanvasTkAgg(fig1, master=inner_frame)
        canvas1.get_tk_widget().pack()
        canvas2 = FigureCanvasTkAgg(fig2, master=inner_frame)
        canvas2.get_tk_widget().pack()
        canvas3 = FigureCanvasTkAgg(fig3, master=inner_frame)
        canvas3.get_tk_widget().pack()

        inner_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox('all'))

    def run_code(self):
        if not self.file_paths or not self.export_dir:
            messagebox.showerror("Error", "Please select files and export directory first.")
            return
        severity= self.severity_level.get()
        # Initialize an empty DataFrame to store the final results
        final_df = pd.DataFrame()

        for file in self.file_paths:
            # Load the CSV files into dataframes
            pca1 = pd.read_excel(file)
            num_columns = len(pca1.columns)
            pca1.columns = ['col' + str(i) for i in range(num_columns)]

            pca1 = pca1.iloc[0:242]
            # Define the columns to select for each new dataframe
            columns = ['col' + str(i) for i in range(1, num_columns)]
            dfs = {col: pca1[['col0', col]] for col in columns}
            # Rename the columns in each DataFrame
            for df in dfs.values():
                df.columns = ['frequency', 'amplitude']

            # Calculate the slope for each DataFrame in dfs
            for df in dfs.values():
                df['slope'] = df['amplitude'].diff()
                df['slope'] = df['slope'].shift(-1)
                df['normalized'] = df['slope']*50 / df['slope'].max()
                df['wcc'] = abs(dfs['col1']['normalized'] - df['normalized'])
                df['frfshift'] = abs(dfs['col1']['amplitude'] - df['amplitude'])
                df['amplituded'] = df['amplitude']
                
            results = {}
            columns = [col for col in pca1.columns if col != 'col0']
            frequency_ranges = [(10, 25), (25, 45),(45,60)]

            for col in columns:
                results[col] = {}
                for i, (low, high) in enumerate(frequency_ranges):
                    mask = (dfs[col]['frequency'] >= low) & (dfs[col]['frequency'] < high)
                    results[col][f'wcc_sum_range_{i+1}'] = dfs[col].loc[mask, 'wcc'].sum()
                    results[col][f'frfshift_sum_range_{i+1}'] = dfs[col].loc[mask, 'frfshift'].sum()
                    results[col][f'peak_frequency_range_{i+1}'] = dfs[col].loc[mask, 'frequency'][dfs[col].loc[mask, 'amplituded'].idxmax()]
                    results[col][f'peak_amplitude_range_{i+1}'] = dfs[col].loc[mask, 'amplituded'].max()

            df_results = pd.DataFrame(results)

            k=4
            # Define the mappings for points and mode shapes
            mode_shapes_mapping = {f'wcc_sum_range_{i}': i for i in range(1, 4)}

            # Initialize an empty DataFrame for high severity
            severity_df = pd.DataFrame(columns=['wcc', 'frfshift', 'severity', 'point', 'mode_shape'])

            start = 2
            combined_df = pd.DataFrame()
            i=0
            while start < len(pca1.columns):
                end = start + k -1
                severity_df = pd.DataFrame()
                for col in pca1.columns[start:end+1]:
                    points_mapping = list(range(1, len(pca1.columns[start:end]) + 2))
                    if i>len(pca1.columns[start:end]):
                        i=0
                    for mode_shape, mode_shape_val in mode_shapes_mapping.items():
                        severity_df = severity_df._append({
                            'wcc': df_results.loc[mode_shape, col],
                            'frfshift': df_results.loc[f'frfshift_sum_range_{mode_shape[-1]}', col],
                            'peak_frequency': df_results.loc[f'peak_frequency_range_{mode_shape[-1]}', col],
                            'peak_amplitude': df_results.loc[f'peak_amplitude_range_{mode_shape[-1]}', col],
                            'severity':  severity,
                            'point': points_mapping[i],
                            'mode_shape': mode_shape_val
                        }, ignore_index=True)
                    i += 1
                combined_df = combined_df._append(severity_df, ignore_index=True)
                start += k

            final_df = pd.concat([final_df, combined_df], axis=0)

        # Save the final DataFrame to an Excel file
        final_df.to_excel(os.path.join(self.export_dir, f'wcc_{severity}.xlsx'), index=False)
        self.display_graphs(final_df)

        messagebox.showinfo("Success", "WCC transformation completed successfully.")

    pass

### Code for PCA Transformation
class PCAGUI:
    def __init__(self, root):
        self.root = root
        self.file_paths = []
        self.export_dir = ''

        self.select_files_button = ctk.CTkButton(root, text="Select Excel Files", command=self.select_files)
        self.select_files_button.pack(fill='both', expand=True, pady=5)

        self.select_export_dir_button = ctk.CTkButton(root, text="Select Destination Folder", command=self.select_export_dir)
        self.select_export_dir_button.pack(fill='both', expand=True, pady=5)

        self.run_button = ctk.CTkButton(root, text="Run", command=self.run_code)
        self.run_button.pack(fill='both', expand=True, pady=5)

    def select_files(self):
        self.file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        print(self.file_paths)

    def select_export_dir(self):
        self.export_dir = filedialog.askdirectory()
        print(self.export_dir)

    def run_code(self):
        if not self.file_paths or not self.export_dir:
            messagebox.showerror("Error", "Please select files and export directory first.")
            return

        for file_path in self.file_paths:
            data = pd.read_excel(file_path)
            num_columns = len(data.columns)
            data.columns = ['col' + str(i) for i in range(num_columns)]
            data = data[(data['col0'] <= 120)]

            data_pca = pd.DataFrame(data.iloc[:, 0])

            for i in range(1, num_columns-1, 5):
                cols = data.columns[i:i+5]
                pca = PCA(n_components=1)
                transformed_data = pca.fit_transform(data[cols])
                pca_data = data[cols].dot(pca.components_.T)
                pca_data = np.ravel(pca_data)
                data_pca[f'col{i//5+1}'] = pca_data

            output_file_path = os.path.join(self.export_dir, f"pca1_{os.path.basename(file_path)}")
            data_pca.to_excel(output_file_path, index=False)

        messagebox.showinfo("Success", "PCA transformation completed successfully.")
    pass


### Code for Combine FRF File
class FRFCombinerGUI():

    def __init__(self, root):

        self.root = root
        self.header = ctk.CTkLabel(root, text="Arrange Excel According to Damage severity (1. Undamaged | 2. High Damaged | 3. Middle Damaged | 4. Low Damaged) ", font=("Arial", 14))
        self.header.pack(pady=10)  # Adjust padding as needed
        self.select_files_button = ctk.CTkButton(root, text="Select Excel Files", command=self.select_files)
        self.select_files_button.pack(fill='both' , expand=True, pady=5)

        self.select_dest_button = ctk.CTkButton(root, text="Select Destination Folder", command=self.select_dest_folder)
        self.select_dest_button.pack(fill='both', expand=True, pady=5)

        self.process_button = ctk.CTkButton(root, text="Run", command=self.process_files)
        self.process_button.pack(fill='both', expand=True, pady=5)

        # Create a progress bar
        self.progress = ttk.Progressbar(root, length=100, mode='indeterminate')
        self.progress.pack(fill='both', pady=10)
        self.progress.pack_forget()

        self.file_paths = []
        self.dest_folder = ""


    def select_files(self):
        
        self.file_paths = filedialog.askopenfilenames(title='Select Excel Files',
                                                      filetypes=(('Excel Files', '*.xlsx'), ('All Files', '*.*')))
        # Sort the file paths
        self.file_paths = natsorted(self.file_paths)

    def select_dest_folder(self):
        self.dest_folder = filedialog.askdirectory(title='Select Destination Folder')

    ### If want to get average frf by combining all sheets
    # def process_files(self):
    #     self.progress.pack(pady=10)
    #     self.progress.start()
    #     # Load the first Excel file
    #     first_file = pd.ExcelFile(self.file_paths[0])

    #     # Get the number of sheets
    #     num_sheets = len(first_file.sheet_names)
    #     all_sheets_df = []

    #     for sheet_num in range(2, num_sheets + 1):
    #         sheet = first_file.sheet_names[sheet_num - 1]
    #         df_list = []
    #         for i, file_path in enumerate(self.file_paths):
    #             df = pd.read_excel(file_path, sheet_name=sheet, nrows=480)
    #             rename_columns(df, str(i))
    #             df_list.append(df)

    #         # Use pd.DataFrame.loc to select columns
    #         dfs = [df.loc[:, ['Frequency'] + [f'{i}_{j}' for j in range(1, 6)]] for i, df in enumerate(df_list)]

    #         # Use pd.DataFrame.merge with on='Frequency' once
    #         final_df = dfs[0]
    #         for df in dfs[1:]:
    #             final_df = final_df.merge(df, on='Frequency')

    #         final_df = final_df[final_df['Frequency'] <= 240]
    #         all_sheets_df.append(final_df)

    #     # Average all sheets
    #     average_df = pd.concat(all_sheets_df).groupby(level=0).mean()

    #     # Export to Excel
    #     average_df.to_excel(f'{self.dest_folder}/average_frf.xlsx', index=False)

    #     self.progress.stop()
    #     self.progress.pack_forget()

    ### If want to get frf data without combining all sheets
    def process_files(self):
        self.progress.pack(pady=10)
        self.progress.start()
        # Load the first Excel file
        first_file = pd.ExcelFile(self.file_paths[0])

        # Get the number of sheets
        num_sheets = len(first_file.sheet_names)

        for sheet_num in range(2, num_sheets + 1):
            sheet = first_file.sheet_names[sheet_num - 1]
            df_list = []
            for i, file_path in enumerate(self.file_paths):
                df = pd.read_excel(file_path, sheet_name=sheet, nrows=480)
                rename_columns(df, str(i))
                df_list.append(df)

            # Use pd.DataFrame.loc to select columns
            dfs = [df.loc[:, ['Frequency'] + [f'{i}_{j}' for j in range(1, 6)]] for i, df in enumerate(df_list)]

            # Use pd.DataFrame.merge with on='Frequency' once
            final_df = dfs[0]
            for df in dfs[1:]:
                final_df = final_df.merge(df, on='Frequency')

            final_df = final_df[final_df['Frequency'] <= 240]
            final_df.to_excel(f'{self.dest_folder}/frf_{sheet}.xlsx', index=False)

            self.progress.stop()
            self.progress.pack_forget()
            
        messagebox.showinfo("Success", "Excel combination completed successfully.")
    pass



### Code for UNV File
class UNVGUI:
    def __init__(self, root):
        self.root = root
        self.file_path = None
        self.dest_folder = None

        self.select_files_button = ctk.CTkButton(root, text="Select Excel Files", command=self.select_files)
        self.select_files_button.pack(fill='both', expand=True, pady=5)

        self.select_dest_button = ctk.CTkButton(root, text="Select Destination Folder", command=self.select_dest_folder)
        self.select_dest_button.pack(fill='both', expand=True, pady=5)

        self.process_button = ctk.CTkButton(root, text="Run", command=self.process_files)
        self.process_button.pack(fill='both', expand=True, pady=5)

    def select_files(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    def select_dest_folder(self):
        self.dest_folder = filedialog.askdirectory()

    def process_files(self):
        if self.file_path is None:
            messagebox.showinfo("Error", "No file selected.")
            return

        if self.dest_folder is None:
            messagebox.showinfo("Error", "No destination folder selected.")
            return

        df = pd.read_excel(self.file_path)

        mode_values = df['Mode'].tolist() #Ex: mode_values = [1,2,3]
        column = len(df.columns[3:13]) // 2 #Ex: column = 10 /2 = 5
        columns = list(range(1, column+1)) #Ex: columns = [1,2,3,4,5]

        save_to_file = 'test_pyuff.UNV'

        #164
        dataset_164 = pyuff.prepare_164(
            units_code=5,
            units_description='SI units',
            temp_mode=1,
            length=1000.0,
            force=1.0,
            temp=1.0,
            temp_offset=1.0)

        #18
        dataset_18_ = d18.generate_text()

        #15
        dataset_15 = pyuff.prepare_15(
            node_nums=[1, 2, 3, 4, 5],
            def_cs=[1, 1, 1, 1, 1],
            disp_cs=[1, 1, 1, 1, 1],
            color=[1, 1, 1, 1, 1],  # I10,
            x=[-210.0, 0.0, 0.0, 0.0, 210.0],
            y=[0.0, 85.0, 0.0, -85.0, 0.0],
            z=[0.0, 0.0, 0.0, 0.0, 0.0])

        #82
        node_pairs = [(1, 2), (2, 5), (4, 5), (1, 4), (1, 3), (3, 5), (2, 3), (3, 4)]

        dataset_82 = []
        for i, pair in enumerate(node_pairs, start=1):
            dataset = pyuff.prepare_82(
                trace_num=i,
                n_nodes=2,
                color=11,
                id='NONE',
                nodes=np.array(pair)
            )
            dataset_82.append(dataset)
        #55

        dataset_55 = []

        modes = mode_values #read list from column mode in excel file, ex: [1,2,3] 
        node_nums = columns #read list column of nodes in excel file, ex: [1,2,3,4,5]
        freqs = list(df['fd (Hz)']) #read list column frequency in excel file, ex: [20.082, 43.079, 51.778]
        damping_ratio = list(df['ζ']) #read list column damping ratio in excel file, ex: [0.04807, 0.0379, 0.04125]
        natural_freq =list(df['ωn (rad/s)'])


        for i, b in enumerate(modes):

            if damping_ratio[i] > 1:
                eig = (-natural_freq[i]*damping_ratio[i]) #if damping ratio > 1, the eigenvalue is real
            else:
                eig = (-natural_freq[i]*damping_ratio[i]) + (np.sqrt((1 - damping_ratio[i] ** 2))*natural_freq[i]*1j) #if damping ratio < 1, the eigenvalue is complex  

            # REMINDER: remove j (complex) for acceleration_realx if data_type = 2
            acceleration_realx = [0.0 + 0j for _ in range(column)] #initialize list of acceleration real x 
            acceleration_realy = [0.0 + 0j for _ in range(column)] #initialize list of acceleration real y
            acceleration_realz = [(df[f'x{m}'].values[i] + df[f'x{m}i'].values[i]*1j) for m in range(1, column+1)] #read list of acceleration real z from excel file

            data = pyuff.prepare_55(
                model_type=1,
                
                id1 = f'Mode Shape: {modes[i]}',
                id2 = f'Natural frequency: {natural_freq[i]}',
                id3 = f'Damping ratio: {damping_ratio[i]}',
                id4 = 'NONE',
                id5 = 'NONE',

                analysis_type=3,
                data_ch=2,
                spec_data_type=12,
                data_type=5,
                eig=eig,

                r1=acceleration_realx,
                r2=acceleration_realy,
                r3=acceleration_realz,
                
                # add r4, r5, r6 if n_data_per_node = 6

                n_data_per_node=3,
                node_nums= columns,
                load_case=1,
                mode_n=i + 1,
                modal_m=0,
                freq=freqs[i],
                modal_damp_vis=0.,
                modal_damp_his=0.)
            
            dataset_55.append(data.copy())


        datasets = [dataset_164, dataset_15]+ dataset_82 + dataset_55

        if save_to_file:
            save_to_file_path = os.path.join(self.dest_folder, save_to_file)
            if os.path.exists(save_to_file_path):
                os.remove(save_to_file_path)
            uffwrite = pyuff.UFF(save_to_file_path)
            for dataset in datasets:
                uffwrite._write_set(dataset, 'add')

        # Append the dataset_18_ text to the file
        with open(save_to_file_path, 'a') as f:
            f.write(dataset_18_)

        messagebox.showinfo("Success", "UNV Transformation completed successfully.")

### Code for Combine FRF File
class FEAFRFCombinerGUI():

    def __init__(self, root):

        self.root = root
        self.header = ctk.CTkLabel(root, text="Arrange CSV According to Damage severity (1. Undamaged | 2. High Damaged | 3. Middle Damaged | 4. Low Damaged) ", font=("Arial", 14))
        self.header.pack(pady=10)  # Adjust padding as needed
        self.select_files_button = ctk.CTkButton(root, text="Select CSV Files", command=self.select_files)
        self.select_files_button.pack(fill='both' , expand=True, pady=5)

        self.select_dest_button = ctk.CTkButton(root, text="Select Destination Folder", command=self.select_dest_folder)
        self.select_dest_button.pack(fill='both', expand=True, pady=5)

        self.process_button = ctk.CTkButton(root, text="Run", command=self.process_files)
        self.process_button.pack(fill='both', expand=True, pady=5)

        # Create a progress bar
        self.progress = ttk.Progressbar(root, length=100, mode='indeterminate')
        self.progress.pack(fill='both', pady=10)
        self.progress.pack_forget()

        self.file_paths = []
        self.dest_folder = ""


    def select_files(self):
        
        self.file_paths = filedialog.askopenfilenames(title='Select CSV Files',
                                                    filetypes=(('CSV Files', '*.csv'), ('All Files', '*.*')))
        # Sort the file paths
        self.file_paths = natsorted(self.file_paths)

    def select_dest_folder(self):
        self.dest_folder = filedialog.askdirectory(title='Select Destination Folder')

    def process_files(self):
        self.progress.pack(pady=10)
        self.progress.start()

        df_list = []
        for i, file_path in enumerate(self.file_paths):
            df = pd.read_csv(file_path, skiprows=3, delimiter=';')
            rename_columns_fea(df, str(i))
            df_list.append(df)

        # Use pd.DataFrame.loc to select columns
        dfs = [df.loc[:, ['Frequency'] + [f'{i}_{j}' for j in range(1, 6)]] for i, df in enumerate(df_list)]

        # Use pd.DataFrame.merge with on='Frequency' once
        final_df = dfs[0]
        for df in dfs[1:]:
            final_df = final_df.merge(df, on='Frequency')

        for col in final_df.columns[1:]:
            final_df[col] = final_df[col] / 15.556
        final_df.to_excel(f'{self.dest_folder}/frf_combined.xlsx', index=False)

        self.progress.stop()
        self.progress.pack_forget()
        messagebox.showinfo("Success", "CSV combination completed successfully.")
    pass


root = ctk.CTk()
main_gui = MainGUI(root)
root.mainloop()



