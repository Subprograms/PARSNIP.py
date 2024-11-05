import os
import subprocess
import tkinter as tk
import pandas as pd
from tkinter import ttk, messagebox
from regipy import RegistryHive
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

# Define file paths (initial paths can be set to the script directory)
sScriptPath = os.path.dirname(os.path.abspath(__file__))
sUsername = os.getlogin()
sNtuserPath = fr"C:\Users\{sUsername}\ntuser.dat"  # Path to actual ntuser.dat

class PARSNIP:
    def __init__(self, root):
        self.root = root
        self.bAutoRefresh = False
        self.xPreviousData = None
        self.sPreviousHiveType = None
        self.sHivePath = ''
        self.nEntryLimit = 100
        self.nInterval = 300

        self.root.title("PARSNIP")
        self.root.geometry("1200x600")

        self.setupUI()

    def setupUI(self):
        """Set up the UI components."""
        xInputFrame = ttk.Frame(self.root)
        xInputFrame.grid(row=0, column=0, columnspan=3, padx=5, pady=20, sticky='ew')

        # Hive Path Input
        ttk.Label(xInputFrame, text="Hive Path:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky='e')
        self.xHivePathInputBox = ttk.Entry(xInputFrame, width=50)
        self.xHivePathInputBox.grid(row=0, column=1, padx=(0, 5), pady=5, sticky='w')
        self.xHivePathSetButton = ttk.Button(xInputFrame, text="Set Path", command=self.setHivePath)
        self.xHivePathSetButton.grid(row=0, column=2, padx=(5, 20), pady=5, sticky='w')

        # Entry Limit Input
        ttk.Label(xInputFrame, text="Entry Limit:").grid(row=0, column=3, padx=(20, 10), pady=5, sticky='e')
        self.xEntryLimitInput = ttk.Entry(xInputFrame, width=10)
        self.xEntryLimitInput.insert(tk.END, str(self.nEntryLimit))
        self.xEntryLimitInput.grid(row=0, column=4, padx=(0, 5), pady=5, sticky='w')
        self.xEntryLimitSetButton = ttk.Button(xInputFrame, text="Set Limit", command=self.setEntryLimit)
        self.xEntryLimitSetButton.grid(row=0, column=5, padx=(5, 20), pady=5, sticky='w')

        # Auto-Refresh Interval Input
        ttk.Label(xInputFrame, text="Auto-Refresh Interval (s):").grid(row=0, column=6, padx=(20, 10), pady=5, sticky='e')
        self.xIntervalInput = ttk.Entry(xInputFrame, width=10)
        self.xIntervalInput.insert(tk.END, str(self.nInterval))
        self.xIntervalInput.grid(row=0, column=7, padx=(0, 5), pady=5, sticky='w')
        self.xIntervalSetButton = ttk.Button(xInputFrame, text="Set Interval", command=self.setInterval)
        self.xIntervalSetButton.grid(row=0, column=8, padx=(5, 0), pady=5, sticky='w')

        # Treeview setup with additional columns
        self.xKeyTrees = ttk.Treeview(
            self.root, 
            columns=('Name', 'Value', 'Type', 'Subkey Count', 'Value Count', 'Key Size', 'Depth'), 
            show='tree headings', 
            selectmode="browse"
        )

        # Set up the headings for all columns
        self.xKeyTrees.heading('#0', text='Key')
        self.xKeyTrees.heading('Name', text='Name')
        self.xKeyTrees.heading('Value', text='Value')
        self.xKeyTrees.heading('Type', text='Type')
        self.xKeyTrees.heading('Subkey Count', text='Subkey Count')
        self.xKeyTrees.heading('Value Count', text='Value Count')
        self.xKeyTrees.heading('Key Size', text='Key Size')
        self.xKeyTrees.heading('Depth', text='Depth')

        # Configure column widths and alignment
        self.xKeyTrees.column('#0', width=250, anchor='center')
        self.xKeyTrees.column('Name', width=150, anchor='center')
        self.xKeyTrees.column('Value', width=300, anchor='center')
        self.xKeyTrees.column('Type', width=100, anchor='center')
        self.xKeyTrees.column('Subkey Count', width=100, anchor='center')
        self.xKeyTrees.column('Value Count', width=100, anchor='center')
        self.xKeyTrees.column('Key Size', width=100, anchor='center')
        self.xKeyTrees.column('Depth', width=100, anchor='center')

        # Scrollbars for Treeview
        xVsb = ttk.Scrollbar(self.root, orient="vertical", command=self.xKeyTrees.yview)
        xHsb = ttk.Scrollbar(self.root, orient="horizontal", command=self.xKeyTrees.xview)
        self.xKeyTrees.configure(yscrollcommand=xVsb.set, xscrollcommand=xHsb.set)

        # Grid layout
        self.xKeyTrees.grid(row=2, column=0, columnspan=3, sticky='nsew')
        xVsb.grid(row=2, column=3, sticky='ns')
        xHsb.grid(row=3, column=0, columnspan=3, sticky='ew')

        # Buttons for manual and auto-refresh
        xAutoRefreshButtonFrame = ttk.Frame(self.root)
        xAutoRefreshButtonFrame.grid(row=4, column=0, columnspan=3, pady=10)

        self.xRefreshButton = ttk.Button(xAutoRefreshButtonFrame, text="Refresh", command=self.refreshPARSNIP)
        self.xRefreshButton.grid(row=0, column=0, padx=5)

        self.xAutoRefreshButton = ttk.Button(xAutoRefreshButtonFrame, text="Enable Auto Refresh", command=self.toggleAutoRefreshPARSNIP)
        self.xAutoRefreshButton.grid(row=0, column=1, padx=5)

        # Frame for changes list
        self.xChangesFrame = ttk.Frame(self.root)
        self.xChangesFrame.grid(row=5, column=0, columnspan=3, sticky='nsew')

        # Treeview for changes list
        self.xChangesList = ttk.Treeview(self.xChangesFrame, columns=('Action', 'Description'), show='headings')
        self.xChangesList.heading('Action', text='Action')
        self.xChangesList.heading('Description', text='Description')

        self.xChangesList.column('Action', width=100, anchor='center')
        self.xChangesList.column('Description', width=800, anchor='w')

        xVsbChanges = ttk.Scrollbar(self.xChangesFrame, orient="vertical", command=self.xChangesList.yview)
        self.xChangesList.configure(yscrollcommand=xVsbChanges.set)

        self.xChangesList.grid(row=0, column=0, sticky='nsew')
        xVsbChanges.grid(row=0, column=1, sticky='ns')

        # Label for loading status
        self.xLoadingLabel = ttk.Label(self.root, text="", anchor='center', font=('Arial', 10, 'italic'))
        self.xLoadingLabel.grid(row=6, column=0, columnspan=3, pady=10, sticky='s')

        # Grid configuration
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.xChangesFrame.grid_rowconfigure(0, weight=1)
        self.xChangesFrame.grid_columnconfigure(0, weight=1)

    def setHivePath(self):
        sInputPath = self.xHivePathInputBox.get().strip()
        self.sHivePath = sInputPath
        messagebox.showinfo("Path Set", f"Hive path set to: {self.sHivePath}")

    def setEntryLimit(self):
        try:
            self.nEntryLimit = int(self.xEntryLimitInput.get().strip())
            messagebox.showinfo("Entry Limit Set", f"Entry limit set to: {self.nEntryLimit}")
        except ValueError:
            messagebox.showerror("Error", "Invalid entry limit. Please enter a valid number.")

    def setInterval(self):
        try:
            self.nInterval = int(self.xIntervalInput.get().strip())
            messagebox.showinfo("Interval Set", f"Auto-refresh interval set to: {self.nInterval} seconds")
        except ValueError:
            messagebox.showerror("Error", "Invalid interval. Please enter a valid number.")

    def parseRegistry(self, sHivePath):
        """Parse Registry using regipy."""
        xData = []
        subkey_counts = {}

        try:
            with ThreadPoolExecutor() as executor:
                xHive = RegistryHive(sHivePath)
                
                for xSubkey in xHive.recurse_subkeys():
                    sKeyPath = xSubkey.path
                    parent_path = '\\'.join(sKeyPath.split('\\')[:-1])
                    if parent_path in subkey_counts:
                        subkey_counts[parent_path] += 1
                    else:
                        subkey_counts[parent_path] = 1
                    
                    nDepth = sKeyPath.count('\\')
                    nKeySize = len(sKeyPath.encode('utf-8'))
                    nValueCount = len(xSubkey.values)
                    nSubkeyCount = subkey_counts.get(sKeyPath, 0)

                    for xValue in xSubkey.values:
                        xData.append({
                            "Key": sKeyPath,
                            "Depth": nDepth,
                            "Key Size": nKeySize,
                            "Subkey Count": nSubkeyCount,
                            "Value Count": nValueCount,
                            "Name": xValue.name,
                            "Value": str(xValue.value),
                            "Type": xValue.value_type
                        })
        except Exception as e:
            messagebox.showerror("Error", f"Error parsing hive: {e}")
        return xData

    def loadGUITrees(self, xData):
        """Load parsed Registry data into Treeview."""
        for xItem in xData:
            xParent = self.xKeyTrees.insert('', 'end', text=xItem['Key'], open=False, tags=('key',))
            self.xKeyTrees.insert(
                xParent, 
                'end', 
                values=(
                    xItem['Name'], 
                    xItem['Value'], 
                    xItem['Type'], 
                    xItem['Subkey Count'], 
                    xItem['Value Count'], 
                    xItem['Key Size'], 
                    xItem['Depth']
                ), 
                tags=('name', 'value', 'type')
            )

    def refreshPARSNIP(self):
        self.xLoadingLabel.config(text="Loading...")
        self.root.update_idletasks()
        
        self.xKeyTrees.delete(*self.xKeyTrees.get_children())
        if self.isLiveHive(self.sHivePath):
            if self.sHivePath.lower() == sNtuserPath.lower():
                self.exportRegistry()
                sParsedPath = self.sExportPath
            else:
                self.exportRegistry()
                sParsedPath = self.sExportPath
        else:
            sParsedPath = self.sHivePath

        if os.path.exists(sParsedPath):
            sHiveType = os.path.basename(sParsedPath).split('.')[0].lower()
            if sHiveType != self.sPreviousHiveType:
                self.xChangesList.delete(*self.xChangesList.get_children())
                self.xPreviousData = None
                self.sPreviousHiveType = sHiveType
            
            xData = self.parseRegistry(sParsedPath)
            self.loadGUITrees(xData)
            
            if self.xPreviousData:
                self.checkChanges(self.xPreviousData, xData)
            
            self.xPreviousData = xData
            self.exportToCSV(xData, 'snapshot')

        self.xLoadingLabel.config(text="")

    def isLiveHive(self, sHivePath):
        sLowercaseHivePath = sHivePath.lower()
        return 'system32' in sLowercaseHivePath or sLowercaseHivePath == sNtuserPath.lower()

    def toggleAutoRefreshPARSNIP(self):
        self.bAutoRefresh = not self.bAutoRefresh
        self.xAutoRefreshButton.config(text="Disable Auto Refresh" if self.bAutoRefresh else "Enable Auto Refresh")
        if self.bAutoRefresh:
            self.autoRefreshPARSNIP()

    def autoRefreshPARSNIP(self):
        if self.bAutoRefresh:
            self.refreshPARSNIP()
            self.root.after(self.nInterval * 1000, self.autoRefreshPARSNIP)

    def exportToCSV(self, xData, sPrefix):
        xDf = pd.DataFrame(xData)
        sTimestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sOutputCsv = os.path.join(sScriptPath, f"{sPrefix}_{sTimestamp}.csv")
        
        if len(xDf) > self.nEntryLimit:
            xDf = xDf.head(self.nEntryLimit)
        
        xDf.to_csv(sOutputCsv, index=False)
        messagebox.showinfo("Export Complete", f"Data exported to: {sOutputCsv}")

# Main function to initialize and run the GUI
def main():
    root = tk.Tk()
    app = PARSNIP(root)
    root.mainloop()

if __name__ == "__main__":
    main()
