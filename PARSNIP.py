import os
import subprocess
import tkinter as tk
import pandas as pd
import numpy as np
import joblib
import re
import chardet
from concurrent.futures import ThreadPoolExecutor
from tkinter import ttk, messagebox
from regipy import RegistryHive
from datetime import datetime

from sklearn.preprocessing import MinMaxScaler, RobustScaler
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split, GridSearchCV, StratifiedKFold
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score, roc_auc_score
from sklearn.feature_selection import RFE

sScriptPath = os.path.dirname(os.path.abspath(__file__))
sUsername = os.getlogin()
sNtuserPath = fr"C:\Users\{sUsername}\ntuser.dat"

class PARSNIP:
    def __init__(self, root):
        self.root = root
        self.bAutoRefresh = False
        self.xPreviousData = None   # Previous snapshot (list of dicts)
        self.sPreviousHiveType = None
        self.sHivePath = ''
        self.nEntryLimit = 100
        self.nInterval = 300
        self.selected_features = None
        self.sRandomForestPath = ''
        self.sClassifiedCsvPath = ''
        self.allData = []  # Current snapshot data
        self.setupUI()

    # Implicit expected columns (defined via a method)
    def get_expected_columns(self):
        return [
            "Key",
            "Depth",
            "Key Size",
            "Subkey Count",
            "Value Count",
            "Label",
            "Tactic",
            "PathCategory_Network Path",
            "PathCategory_Other Path",
            "PathCategory_Service Path",
            "PathCategory_Startup Path",
            "TypeGroup_Binary",
            "TypeGroup_Numeric",
            "TypeGroup_Others",
            "TypeGroup_String",
            "KeyNameCategory_Other Keys",
            "KeyNameCategory_Run Keys",
            "KeyNameCategory_Security and Configuration Keys",
            "KeyNameCategory_Service Keys",
            "KeyNameCategory_Internet and Network Keys",
            "Value Processed",
            "Name",
            "Value",
            "Type"
        ]

    def unifyFinalColumns(self, df):
        expected = self.get_expected_columns()
        for col in expected:
            if col not in df.columns:
                df.loc[:, col] = 0
        df = df[expected]
        return df

    def categorizePath(self, p):
        if "Run" in p:
            return "Startup Path"
        elif "Services" in p:
            return "Service Path"
        elif "Internet Settings" in p:
            return "Network Path"
        return "Other Path"

    def mapType(self, t):
        type_map = {
            "String": ["REG_SZ", "REG_EXPAND_SZ", "REG_MULTI_SZ"],
            "Numeric": ["REG_DWORD", "REG_QWORD"],
            "Binary": ["REG_BINARY"],
            "Others": ["REG_NONE", "REG_LINK", "0"]
        }
        for group, vals in type_map.items():
            if t in vals:
                return group
        return "Others"

    def categorizeKeyName(self, kn):
        categories = {
            "Run Keys": ["Run", "RunOnce", "RunServices"],
            "Service Keys": ["ImageFileExecutionOptions", "AppInit_DLLs"],
            "Security and Configuration Keys": ["Policies", "Explorer"],
            "Internet and Network Keys": ["ProxyEnable", "ProxyServer"],
            "File Execution Keys": ["ShellExecuteHooks"]
        }
        low = kn.lower()
        for cat, keys in categories.items():
            if any(k.lower() in low for k in keys):
                return cat
        return "Other Keys"

    def preprocessValue(self, v):
        if isinstance(v, str):
            return len(v)
        return v

    def preprocessData(self, df):
        if df.empty:
            return pd.DataFrame()
        xDf = df.copy()
        numeric_df = xDf.select_dtypes(include=[np.number])
        xDf.fillna(numeric_df.mean(), inplace=True)

        # Create one-hot encoded columns for Path, Type, and Key Name
        xDf['Path Category'] = xDf['Key'].apply(self.categorizePath)
        path_dummies = pd.get_dummies(xDf['Path Category'], prefix='PathCategory')
        xDf = pd.concat([xDf, path_dummies], axis=1)

        xDf['Type Group'] = xDf['Type'].apply(self.mapType)
        type_dummies = pd.get_dummies(xDf['Type Group'], prefix='TypeGroup')
        xDf = pd.concat([xDf, type_dummies], axis=1)

        xDf['Key Name Category'] = xDf['Name'].apply(self.categorizeKeyName)
        name_dummies = pd.get_dummies(xDf['Key Name Category'], prefix='KeyNameCategory')
        xDf = pd.concat([xDf, name_dummies], axis=1)

        xDf['Value Processed'] = xDf['Value'].apply(self.preprocessValue)

        # Scale numeric columns
        scaler_minmax = MinMaxScaler()
        for col in ['Depth', 'Value Count', 'Value Processed']:
            if col in xDf.columns:
                xDf[[col]] = scaler_minmax.fit_transform(xDf[[col]])
        scaler_robust = RobustScaler()
        for col in ['Key Size', 'Subkey Count']:
            if col in xDf.columns:
                xDf[[col]] = scaler_robust.fit_transform(xDf[[col]])

        return xDf

    def preprocessAndExport(self, xData):
        if not xData:
            return
        xDf = pd.DataFrame(xData)
        preproc = self.preprocessData(xDf)
        preproc = self.unifyFinalColumns(preproc)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_csv = os.path.join(sScriptPath, f"preprocessed_{ts}.csv")
        if len(preproc) > self.nEntryLimit:
            preproc = preproc.head(self.nEntryLimit)
        preproc.to_csv(out_csv, index=False)
        
    def setupUI(self):
        # Input frame for basic parameters
        xInputFrame = ttk.Frame(self.root)
        xInputFrame.grid(row=0, column=0, columnspan=3, padx=5, pady=20, sticky='ew')
        ttk.Label(xInputFrame, text="Hive Path:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky='e')
        self.xHivePathInputBox = ttk.Entry(xInputFrame, width=50)
        self.xHivePathInputBox.grid(row=0, column=1, padx=(0, 5), pady=5, sticky='w')
        self.xHivePathSetButton = ttk.Button(xInputFrame, text="Set Path", command=self.setHivePath)
        self.xHivePathSetButton.grid(row=0, column=2, padx=(5, 20), pady=5, sticky='w')
        ttk.Label(xInputFrame, text="Entry Limit:").grid(row=0, column=3, padx=(20, 10), pady=5, sticky='e')
        self.xEntryLimitInput = ttk.Entry(xInputFrame, width=10)
        self.xEntryLimitInput.insert(tk.END, str(self.nEntryLimit))
        self.xEntryLimitInput.grid(row=0, column=4, padx=(0, 5), pady=5, sticky='w')
        self.xEntryLimitSetButton = ttk.Button(xInputFrame, text="Set Limit", command=self.setEntryLimit)
        self.xEntryLimitSetButton.grid(row=0, column=5, padx=(5, 20), pady=5, sticky='w')
        ttk.Label(xInputFrame, text="Auto-Refresh Interval (s):").grid(row=0, column=6, padx=(20, 10), pady=5, sticky='e')
        self.xIntervalInput = ttk.Entry(xInputFrame, width=10)
        self.xIntervalInput.insert(tk.END, str(self.nInterval))
        self.xIntervalInput.grid(row=0, column=7, padx=(0, 5), pady=5, sticky='w')
        self.xIntervalSetButton = ttk.Button(xInputFrame, text="Set Interval", command=self.setInterval)
        self.xIntervalSetButton.grid(row=0, column=8, padx=(5, 0), pady=5, sticky='w')
        # Frame for ML file inputs
        xMLFrame = ttk.Frame(self.root)
        xMLFrame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
        ttk.Label(xMLFrame, text="Random Forest Model File:").grid(row=0, column=0, padx=(0,10), pady=5, sticky='e')
        self.xRFInput = ttk.Entry(xMLFrame, width=50)
        self.xRFInput.grid(row=0, column=1, padx=(0,5), pady=5, sticky='w')
        self.xRFSetButton = ttk.Button(xMLFrame, text="Set RF File", command=self.setRandomForestPath)
        self.xRFSetButton.grid(row=0, column=2, padx=(5,20), pady=5, sticky='w')
        ttk.Label(xMLFrame, text="Classified Changes CSV:").grid(row=0, column=3, padx=(20,10), pady=5, sticky='e')
        self.xClassCsvInput = ttk.Entry(xMLFrame, width=50)
        self.xClassCsvInput.grid(row=0, column=4, padx=(0,5), pady=5, sticky='w')
        self.xClassCsvSetButton = ttk.Button(xMLFrame, text="Set CSV", command=self.setClassifiedCsvPath)
        self.xClassCsvSetButton.grid(row=0, column=5, padx=(5,0), pady=5, sticky='w')
        # Treeview for displaying registry keys
        self.xKeyTrees = ttk.Treeview(self.root, columns=('Name', 'Value', 'Type', 'Subkey Count', 'Value Count', 'Key Size', 'Depth'), show='tree headings', selectmode="browse")
        self.xKeyTrees.heading('#0', text='Key', command=lambda: self.sortTreeview('#0', False))
        self.xKeyTrees.heading('Name', text='Name', command=lambda: self.sortTreeview('Name', False))
        self.xKeyTrees.heading('Value', text='Value', command=lambda: self.sortTreeview('Value', False))
        self.xKeyTrees.heading('Type', text='Type', command=lambda: self.sortTreeview('Type', False))
        self.xKeyTrees.heading('Subkey Count', text='Subkey Count', command=lambda: self.sortTreeview('Subkey Count', False))
        self.xKeyTrees.heading('Value Count', text='Value Count', command=lambda: self.sortTreeview('Value Count', False))
        self.xKeyTrees.heading('Key Size', text='Key Size', command=lambda: self.sortTreeview('Key Size', False))
        self.xKeyTrees.heading('Depth', text='Depth', command=lambda: self.sortTreeview('Depth', False))
        self.xKeyTrees.column('#0', width=250, anchor='center')
        self.xKeyTrees.column('Name', width=150, anchor='center')
        self.xKeyTrees.column('Value', width=300, anchor='center')
        self.xKeyTrees.column('Type', width=100, anchor='center')
        self.xKeyTrees.column('Subkey Count', width=100, anchor='center')
        self.xKeyTrees.column('Value Count', width=100, anchor='center')
        self.xKeyTrees.column('Key Size', width=100, anchor='center')
        self.xKeyTrees.column('Depth', width=100, anchor='center')
        xVsb = ttk.Scrollbar(self.root, orient="vertical", command=self.xKeyTrees.yview)
        xHsb = ttk.Scrollbar(self.root, orient="horizontal", command=self.xKeyTrees.xview)
        self.xKeyTrees.configure(yscrollcommand=xVsb.set, xscrollcommand=xHsb.set)
        self.xKeyTrees.tag_configure('key', background='lightblue')
        self.xKeyTrees.tag_configure('name', background='lightgreen')
        self.xKeyTrees.tag_configure('value', background='lightyellow')
        self.xKeyTrees.tag_configure('type', background='lightpink')
        self.xKeyTrees.grid(row=2, column=0, columnspan=3, sticky='nsew')
        xVsb.grid(row=2, column=3, sticky='ns')
        xHsb.grid(row=3, column=0, columnspan=3, sticky='ew')
        # Frame for auto-refresh buttons
        xAutoRefreshButtonFrame = ttk.Frame(self.root)
        xAutoRefreshButtonFrame.grid(row=4, column=0, columnspan=3, pady=10)
        self.xRefreshButton = ttk.Button(xAutoRefreshButtonFrame, text="Refresh", command=self.refreshPARSNIP)
        self.xRefreshButton.grid(row=0, column=0, padx=5)
        self.xAutoRefreshButton = ttk.Button(xAutoRefreshButtonFrame, text="Enable Auto Refresh", command=self.toggleAutoRefreshPARSNIP)
        self.xAutoRefreshButton.grid(row=0, column=1, padx=5)
        # Search frame for filtering keys
        xSearchFrame = ttk.Frame(self.root)
        xSearchFrame.grid(row=5, column=0, columnspan=3, pady=5, sticky='ew')
        ttk.Label(xSearchFrame, text="Search Keyword:").grid(row=0, column=0, padx=(0,10), pady=5, sticky='e')
        self.xSearchInput = ttk.Entry(xSearchFrame, width=30)
        self.xSearchInput.grid(row=0, column=1, padx=(0,5), pady=5, sticky='w')
        self.xSearchButton = ttk.Button(xSearchFrame, text="Search", command=self.searchKeys)
        self.xSearchButton.grid(row=0, column=2, padx=5, pady=5)
        self.xClearSearchButton = ttk.Button(xSearchFrame, text="Clear Search", command=self.clearSearch)
        self.xClearSearchButton.grid(row=0, column=3, padx=5, pady=5)
        # Frame for changes list
        self.xChangesFrame = ttk.Frame(self.root)
        self.xChangesFrame.grid(row=6, column=2, columnspan=3, sticky='nsew')
        self.xChangesList = ttk.Treeview(self.xChangesFrame, columns=('Action', 'Description'), show='headings')
        self.xChangesList.heading('Action', text='Action')
        self.xChangesList.heading('Description', text='Description')
        self.xChangesList.column('Action', width=300, anchor='center')
        self.xChangesList.column('Description', width=900, anchor='w')
        xVsbChanges = ttk.Scrollbar(self.xChangesFrame, orient="vertical", command=self.xChangesList.yview)
        self.xChangesList.configure(yscrollcommand=xVsbChanges.set)
        self.xChangesList.grid(row=0, column=0, sticky='nsew')
        xVsbChanges.grid(row=0, column=1, sticky='ns')
        # Configure tags for changes: red for Malicious, green for Benign
        self.xChangesList.tag_configure("Malicious", background="lightpink")
        self.xChangesList.tag_configure("Benign", background="lightgreen")
        # Loading Label
        self.xLoadingLabel = ttk.Label(self.root, text="", anchor='center', font=('Arial', 10, 'italic'))
        self.xLoadingLabel.grid(row=7, column=0, columnspan=3, pady=10, sticky='s')

    def setHivePath(self):
        path = self.xHivePathInputBox.get().strip()
        self.sHivePath = path
        messagebox.showinfo("Path Set", f"Hive path set to: {path}")

    def setEntryLimit(self):
        try:
            self.nEntryLimit = int(self.xEntryLimitInput.get().strip())
            messagebox.showinfo("Entry Limit Set", f"Entry limit set to: {self.nEntryLimit}")
        except ValueError:
            messagebox.showerror("Error", "Invalid entry limit.")

    def setInterval(self):
        try:
            self.nInterval = int(self.xIntervalInput.get().strip())
            messagebox.showinfo("Interval Set", f"Auto-refresh interval set to: {self.nInterval} seconds")
        except ValueError:
            messagebox.showerror("Error", "Invalid interval.")

    def setRandomForestPath(self):
        rfpath = self.xRFInput.get().strip()
        self.sRandomForestPath = rfpath
        messagebox.showinfo("Path Set", f"Random Forest file set to: {rfpath}")

    def setClassifiedCsvPath(self):
        csvpath = self.xClassCsvInput.get().strip()
        self.sClassifiedCsvPath = csvpath
        messagebox.showinfo("Path Set", f"Classified CSV set to: {csvpath}")

    # Registry parsing
    def parseRegistry(self, hive_path):
        xData = []
        subkey_counts = {}
        try:
            with ThreadPoolExecutor() as executor:
                hive = RegistryHive(hive_path)
                for subkey in hive.recurse_subkeys():
                    kpath = subkey.path
                    parent_path = '\\'.join(kpath.split('\\')[:-1])
                    subkey_counts[parent_path] = subkey_counts.get(parent_path, 0) + 1
                    depth = kpath.count('\\')
                    ksize = len(kpath.encode('utf-8'))
                    vcount = len(subkey.values)
                    scount = subkey_counts.get(kpath, 0)
                    for val in subkey.values:
                        xData.append({
                            "Key": kpath,
                            "Depth": depth,
                            "Key Size": ksize,
                            "Subkey Count": scount,
                            "Value Count": vcount,
                            "Name": str(val.name) if val.name else "0",
                            "Value": str(val.value) if val.value else "0",
                            "Type": str(val.value_type) if val.value_type else "0"
                        })
            self.preprocessAndExport(xData)
        except Exception as e:
            messagebox.showerror("Error", f"Error parsing hive: {e}")
        return xData

    # Classification of changes
    def classifyChanges(self, changes_df):
        if changes_df.empty:
            return changes_df

        actions = changes_df["Action"] if "Action" in changes_df.columns else None
        proc_df = self.preprocessData(changes_df)
        df_unified = self.unifyFinalColumns(proc_df).copy()

        # Add the preserved "Action" column back into the unified DataFrame for refernece
        if actions is not None:
            df_unified["Action"] = actions.values

        if not self.sRandomForestPath or not os.path.exists(self.sRandomForestPath):
            messagebox.showerror("Error", "Random Forest model file not set or not found.")
            df_unified.loc[:, 'Predicted Label'] = "Unclassified"
            df_unified.loc[:, 'Change Detected Datetime'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            return df_unified

        model = joblib.load(self.sRandomForestPath)
        if hasattr(model, "feature_names_in_"):
            selected = model.feature_names_in_
            try:
                X = df_unified.loc[:, selected].copy()
            except KeyError:
                messagebox.showerror("Error", "Some expected feature columns are missing in the input data.")
                df_unified.loc[:, 'Predicted Label'] = "Unclassified"
                df_unified.loc[:, 'Change Detected Datetime'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                return df_unified
        else:
            non_features = ['Key', 'Name', 'Value', 'Type', 'Path Category', 'Type Group', 'Key Name Category', 'Label', 'Tactic']
            X = df_unified.drop(columns=non_features, errors='ignore').copy()

        for col in X.columns:
            X[col] = pd.to_numeric(X[col], errors='coerce').fillna(0)

        y_scores = model.predict_proba(X)[:, 1]
        df_unified.loc[:, 'Predicted Label'] = np.where(y_scores >= 0.5, 'Malicious', 'Benign')
        df_unified.loc[:, 'Change Detected Datetime'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return df_unified

    def appendClassifiedCsv(self, df):
        if df.empty:
            return
        if self.sClassifiedCsvPath and os.path.exists(self.sClassifiedCsvPath):
            try:
                existing = pd.read_csv(self.sClassifiedCsvPath, dtype=str)
                combined = pd.concat([existing, df], ignore_index=True)
                combined.to_csv(self.sClassifiedCsvPath, index=False)
            except Exception:
                df.to_csv(self.sClassifiedCsvPath, index=False)
        else:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_csv = os.path.join(sScriptPath, f"classified_changes_{ts}.csv")
            df.to_csv(new_csv, index=False)
            self.sClassifiedCsvPath = new_csv
            self.xClassCsvInput.delete(0, tk.END)
            self.xClassCsvInput.insert(0, new_csv)
            messagebox.showinfo("CSV Created", f"Classified CSV created at: {new_csv}")

    # Load data into the treeview
    def loadGUITrees(self, xData):
        self.xKeyTrees.delete(*self.xKeyTrees.get_children())
        for entry in xData:
            # Insert the parent with the "key" tag to color it blue
            parent = self.xKeyTrees.insert('', 'end', text=entry['Key'], open=True, tags=('key',))
            self.xKeyTrees.insert(parent, 'end', values=(
                entry.get('Name', ''),
                entry.get('Value', ''),
                entry.get('Type', ''),
                entry.get('Subkey Count', ''),
                entry.get('Value Count', ''),
                entry.get('Key Size', ''),
                entry.get('Depth', '')
            ))

    # Change detection: compare previous and current snapshot
    def checkChanges(self, previous, current):
        changes = []
        unmatched_prev = previous.copy()

        # Compare each current entry against previous entries using a composite key
        for curr in current:
            # Create a composite key tuple for the current entry
            curr_key = (curr.get("Key"), curr.get("Name"), curr.get("Type"))
            found_match = None

            # Search for a matching entry in the previous snapshot
            for prev in unmatched_prev:
                prev_key = (prev.get("Key"), prev.get("Name"), prev.get("Type"))
                if curr_key == prev_key:
                    found_match = prev
                    break

            if found_match:
                # Remove the matched previous entry so it is not processed again
                unmatched_prev.remove(found_match)
                # If any field differs, mark as modified
                if curr != found_match:
                    entry = curr.copy()
                    entry['Action'] = 'Modified'
                    changes.append(entry)
            else:
                # No matching previous entry found means this entry was added
                entry = curr.copy()
                entry['Action'] = 'Added'
                changes.append(entry)

        # Any entries left in unmatched_prev were removed
        for prev in unmatched_prev:
            entry = prev.copy()
            entry['Action'] = 'Removed'
            changes.append(entry)

        return pd.DataFrame(changes)

    # Main refresh: parse hive, detect changes, classify, and update UI
    def refreshPARSNIP(self):
        self.xLoadingLabel.config(text="Loading...")
        self.root.update_idletasks()

        if self.isLiveHive(self.sHivePath):
            self.exportRegistry()
            parsed_path = getattr(self, 'sExportPath', self.sHivePath)
        else:
            parsed_path = self.sHivePath

        if os.path.exists(parsed_path):
            # Get the base hive name and normalize it by removing trailing " (number)"
            base_hive = os.path.basename(parsed_path).split('.')[0].lower()
            hive_type = re.sub(r'\s*\(\d+\)$', '', base_hive)
            if hive_type != self.sPreviousHiveType:
                self.xChangesList.delete(*self.xChangesList.get_children())
                self.xPreviousData = None
                self.sPreviousHiveType = hive_type

            current_data = self.parseRegistry(parsed_path)
            self.allData = current_data
            self.loadGUITrees(current_data)

            if self.xPreviousData is not None:
                changes = self.checkChanges(self.xPreviousData, current_data)
                if not changes.empty:
                    classified = self.classifyChanges(changes)
                    self.appendClassifiedCsv(classified)
                    for idx, row in classified.iterrows():
                        tag = "Malicious" if row["Predicted Label"] == "Malicious" else "Benign"
                        self.xChangesList.insert("", "end", values=(
                            row["Action"],
                            f"{row['Key']} => {row['Predicted Label']} at {row['Change Detected Datetime']}"
                        ), tags=(tag,))
            self.xPreviousData = current_data
            self.exportToCSV(current_data, 'snapshot')
        self.xLoadingLabel.config(text="")

    # CSV Export helpers
    def exportToCSV(self, data, prefix):
        columns = ["Key", "Name", "Value", "Type", "Subkey Count", "Value Count", "Key Size", "Depth"]
        df = pd.DataFrame(data, columns=columns)
        df.dropna(axis=1, how='all', inplace=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_csv = os.path.join(sScriptPath, f"{prefix}_{ts}.csv")
        if len(df) > self.nEntryLimit:
            df = df.head(self.nEntryLimit)
        df.to_csv(out_csv, index=False)
        messagebox.showinfo("Export Complete", f"Data exported to: {out_csv}")

    def exportSortedCSV(self):
        columns = ["Key", "Name", "Value", "Type", "Subkey Count", "Value Count", "Key Size", "Depth"]
        data = []
        for parent in self.xKeyTrees.get_children(''):
            key_text = self.xKeyTrees.item(parent, 'text')
            for child in self.xKeyTrees.get_children(parent):
                data.append({
                    "Key": key_text,
                    "Name": self.xKeyTrees.set(child, 'Name'),
                    "Value": self.xKeyTrees.set(child, 'Value'),
                    "Type": self.xKeyTrees.set(child, 'Type'),
                    "Subkey Count": self.xKeyTrees.set(child, 'Subkey Count'),
                    "Value Count": self.xKeyTrees.set(child, 'Value Count'),
                    "Key Size": self.xKeyTrees.set(child, 'Key Size'),
                    "Depth": self.xKeyTrees.set(child, 'Depth')
                })
        df = pd.DataFrame(data, columns=columns)
        df.dropna(axis=1, how='all', inplace=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_csv = os.path.join(sScriptPath, f"snapshot_sorted_{ts}.csv")
        df.to_csv(out_csv, index=False)
        messagebox.showinfo("Export Complete", f"Sorted data exported to: {out_csv}")

    def sortTreeview(self, col, reverse):
        items = []
        if col == '#0':
            for it in self.xKeyTrees.get_children(''):
                key_text = self.xKeyTrees.item(it, 'text')
                items.append((key_text, it))
            sorted_items = sorted(items, key=lambda x: x[0].lower(), reverse=reverse)
            for idx, itm in enumerate(sorted_items):
                self.xKeyTrees.move(itm[1], '', idx)
        else:
            for parent in self.xKeyTrees.get_children(''):
                parent_text = self.xKeyTrees.item(parent, 'text')
                child_vals = [self.xKeyTrees.set(child, col) for child in self.xKeyTrees.get_children(parent)]
                items.append((parent_text, child_vals, parent))
            sorted_items = sorted(items, key=lambda x: x[1][0], reverse=reverse)
            for idx, itm in enumerate(sorted_items):
                self.xKeyTrees.move(itm[2], '', idx)
        self.xKeyTrees.heading(col, command=lambda: self.sortTreeview(col, not reverse))
        self.exportSortedCSV()

    def toggleAutoRefreshPARSNIP(self):
        self.bAutoRefresh = not self.bAutoRefresh
        self.xAutoRefreshButton.config(text="Disable Auto Refresh" if self.bAutoRefresh else "Enable Auto Refresh")
        if self.bAutoRefresh:
            self.autoRefreshPARSNIP()

    def autoRefreshPARSNIP(self):
        if self.bAutoRefresh:
            self.refreshPARSNIP()
            self.root.after(self.nInterval * 1000, self.autoRefreshPARSNIP)

    # Setters for UI Inputs
    def setHivePath(self):
        path = self.xHivePathInputBox.get().strip()
        self.sHivePath = path
        messagebox.showinfo("Path Set", f"Hive path set to: {path}")

    def setEntryLimit(self):
        try:
            self.nEntryLimit = int(self.xEntryLimitInput.get().strip())
            messagebox.showinfo("Entry Limit Set", f"Entry limit set to: {self.nEntryLimit}")
        except ValueError:
            messagebox.showerror("Error", "Invalid entry limit.")

    def setInterval(self):
        try:
            self.nInterval = int(self.xIntervalInput.get().strip())
            messagebox.showinfo("Interval Set", f"Auto-refresh interval set to: {self.nInterval} seconds")
        except ValueError:
            messagebox.showerror("Error", "Invalid interval.")

    def setRandomForestPath(self):
        rfpath = self.xRFInput.get().strip()
        self.sRandomForestPath = rfpath
        messagebox.showinfo("Path Set", f"Random Forest file set to: {rfpath}")

    def setClassifiedCsvPath(self):
        csvpath = self.xClassCsvInput.get().strip()
        self.sClassifiedCsvPath = csvpath
        messagebox.showinfo("Path Set", f"Classified CSV set to: {csvpath}")

    # Registry Export (Live Hive)
    def exportRegistry(self):
        hive_type = os.path.basename(self.sHivePath).split('.')[0].lower()
        hive_ext = os.path.splitext(self.sHivePath)[1]
        if hive_type == 'ntuser':
            hive_param = 'HKCU'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath))
        elif hive_type == 'system':
            hive_param = 'HKLM\\System'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + hive_ext)
        elif hive_type == 'software':
            hive_param = 'HKLM\\Software'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + hive_ext)
        elif hive_type == 'sam':
            hive_param = 'HKLM\\SAM'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + hive_ext)
        elif hive_type == 'security':
            hive_param = 'HKLM\\SECURITY'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + hive_ext)
        elif hive_type == 'hardware':
            hive_param = 'HKLM\\HARDWARE'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + hive_ext)
        else:
            hive_param = 'HKLM'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + hive_ext)
        try:
            subprocess.check_call(['reg', 'save', hive_param, self.sExportPath, '/y'])
            messagebox.showinfo("Success", f"Unparsed Registry exported to: {self.sExportPath}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Error exporting registry: {e}")

    def isLiveHive(self, path):
        low = path.lower()
        return ('system32' in low) or (low == sNtuserPath.lower())

    def searchKeys(self):
        kw = self.xSearchInput.get().strip().lower()
        if not self.allData:
            messagebox.showinfo("Search", "No data loaded to search.")
            return
        filtered = [row for row in self.allData if kw in row['Key'].lower()]
        self.loadGUITrees(filtered)

    def clearSearch(self):
        self.xSearchInput.delete(0, tk.END)
        self.loadGUITrees(self.allData)

    def exportToCSV(self, data, prefix):
        columns = ["Key", "Name", "Value", "Type", "Subkey Count", "Value Count", "Key Size", "Depth"]
        df = pd.DataFrame(data, columns=columns)
        df.dropna(axis=1, how='all', inplace=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_csv = os.path.join(sScriptPath, f"{prefix}_{ts}.csv")
        if len(df) > self.nEntryLimit:
            df = df.head(self.nEntryLimit)
        df.to_csv(out_csv, index=False)
        messagebox.showinfo("Export Complete", f"Data exported to: {out_csv}")

    def exportToCSV(self, data, prefix):
        columns = ["Key", "Name", "Value", "Type", "Subkey Count", "Value Count", "Key Size", "Depth"]
        df = pd.DataFrame(data, columns=columns)
        df.dropna(axis=1, how='all', inplace=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_csv = os.path.join(sScriptPath, f"{prefix}_{ts}.csv")
        if len(df) > self.nEntryLimit:
            df = df.head(self.nEntryLimit)
        df.to_csv(out_csv, index=False)
        messagebox.showinfo("Export Complete", f"Data exported to: {out_csv}")

    def sortTreeview(self, col, reverse):
        items = []
        if col == '#0':
            for it in self.xKeyTrees.get_children(''):
                key_text = self.xKeyTrees.item(it, 'text')
                items.append((key_text, it))
            sorted_items = sorted(items, key=lambda x: x[0].lower(), reverse=reverse)
            for idx, itm in enumerate(sorted_items):
                self.xKeyTrees.move(itm[1], '', idx)
        else:
            for parent in self.xKeyTrees.get_children(''):
                parent_text = self.xKeyTrees.item(parent, 'text')
                child_vals = [self.xKeyTrees.set(child, col) for child in self.xKeyTrees.get_children(parent)]
                items.append((parent_text, child_vals, parent))
            sorted_items = sorted(items, key=lambda x: x[1][0], reverse=reverse)
            for idx, itm in enumerate(sorted_items):
                self.xKeyTrees.move(itm[2], '', idx)
        self.xKeyTrees.heading(col, command=lambda: self.sortTreeview(col, not reverse))
        self.exportSortedCSV()

    def toggleAutoRefreshPARSNIP(self):
        self.bAutoRefresh = not self.bAutoRefresh
        self.xAutoRefreshButton.config(text="Disable Auto Refresh" if self.bAutoRefresh else "Enable Auto Refresh")
        if self.bAutoRefresh:
            self.autoRefreshPARSNIP()

    def autoRefreshPARSNIP(self):
        if self.bAutoRefresh:
            self.refreshPARSNIP()
            self.root.after(self.nInterval * 1000, self.autoRefreshPARSNIP)

def main():
    root = tk.Tk()
    app = PARSNIP(root)
    root.mainloop()

if __name__ == "__main__":
    main()
