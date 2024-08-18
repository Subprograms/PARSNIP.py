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
        self.sPreviousHiveType = None  # Track the previous hive type
        self.sHivePath = ''  # Initialize sHivePath as an instance variable
        self.entry_limit = 100  # Default entry limit
        self.dInterval = 300  # Default interval in seconds

        self.root.title("PARSNIP")
        self.root.geometry("1200x600")

        self.setupUI()

    def setupUI(self):
        """Set up the UI components."""
        inputFrame = ttk.Frame(self.root)
        inputFrame.grid(row=0, column=0, columnspan=3, padx=5, pady=20, sticky='ew')

        # Hive Path Input
        ttk.Label(inputFrame, text="Hive Path:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky='e')
        self.sHivePathInputBox = ttk.Entry(inputFrame, width=50)
        self.sHivePathInputBox.insert(tk.END, sScriptPath + '\\<hive>')  # Initial value set to script path
        self.sHivePathInputBox.grid(row=0, column=1, padx=(0, 5), pady=5, sticky='w')
        self.sHivePathSetButton = ttk.Button(inputFrame, text="Set Path", command=self.setHivePath)
        self.sHivePathSetButton.grid(row=0, column=2, padx=(5, 20), pady=5, sticky='w')

        # Entry Limit Input
        ttk.Label(inputFrame, text="Entry Limit:").grid(row=0, column=3, padx=(20, 10), pady=5, sticky='e')
        self.entryLimitInput = ttk.Entry(inputFrame, width=10)
        self.entryLimitInput.insert(tk.END, str(self.entry_limit))
        self.entryLimitInput.grid(row=0, column=4, padx=(0, 5), pady=5, sticky='w')
        self.entryLimitSetButton = ttk.Button(inputFrame, text="Set Limit", command=self.setEntryLimit)
        self.entryLimitSetButton.grid(row=0, column=5, padx=(5, 20), pady=5, sticky='w')

        # Auto-Refresh Interval Input
        ttk.Label(inputFrame, text="Auto-Refresh Interval (s):").grid(row=0, column=6, padx=(20, 10), pady=5, sticky='e')
        self.intervalInput = ttk.Entry(inputFrame, width=10)
        self.intervalInput.insert(tk.END, str(self.dInterval))
        self.intervalInput.grid(row=0, column=7, padx=(0, 5), pady=5, sticky='w')
        self.intervalSetButton = ttk.Button(inputFrame, text="Set Interval", command=self.setInterval)
        self.intervalSetButton.grid(row=0, column=8, padx=(5, 0), pady=5, sticky='w')

        inputFrame.grid_columnconfigure(0, weight=1)
        inputFrame.grid_columnconfigure(1, weight=1)
        inputFrame.grid_columnconfigure(2, weight=1)
        inputFrame.grid_columnconfigure(3, weight=1)
        inputFrame.grid_columnconfigure(4, weight=1)
        inputFrame.grid_columnconfigure(5, weight=1)
        inputFrame.grid_columnconfigure(6, weight=1)
        inputFrame.grid_columnconfigure(7, weight=1)
        inputFrame.grid_columnconfigure(8, weight=1)

        # Treeview setup
        self.KeyTrees = ttk.Treeview(self.root, columns=('Name', 'Value', 'Type'), show='tree headings', selectmode="browse")
        self.KeyTrees.heading('#0', text='Key', command=lambda: self.sortTreeview('#0', False))
        self.KeyTrees.heading('Name', text='Name', command=lambda: self.sortTreeview('Name', False))
        self.KeyTrees.heading('Value', text='Value', command=lambda: self.sortTreeview('Value', False))
        self.KeyTrees.heading('Type', text='Type', command=lambda: self.sortTreeview('Type', False))

        # Treeview columns configuration
        self.KeyTrees.column('#0', width=250, anchor='center')
        self.KeyTrees.column('Name', width=150, anchor='center')
        self.KeyTrees.column('Value', width=300, anchor='center')
        self.KeyTrees.column('Type', width=100, anchor='center')

        # Scrollbars for Treeview
        vsb = ttk.Scrollbar(self.root, orient="vertical", command=self.KeyTrees.yview)
        hsb = ttk.Scrollbar(self.root, orient="horizontal", command=self.KeyTrees.xview)
        self.KeyTrees.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Tags for Treeview items
        self.KeyTrees.tag_configure('key', background='lightblue')
        self.KeyTrees.tag_configure('name', background='lightgreen')
        self.KeyTrees.tag_configure('value', background='lightyellow')
        self.KeyTrees.tag_configure('type', background='lightpink')

        # Grid layout
        self.KeyTrees.grid(row=2, column=0, columnspan=3, sticky='nsew')
        vsb.grid(row=2, column=3, sticky='ns')
        hsb.grid(row=3, column=0, columnspan=3, sticky='ew')

        # Buttons for manual and auto-refresh
        bAutoRefreshButtonFrame = ttk.Frame(self.root)
        bAutoRefreshButtonFrame.grid(row=4, column=0, columnspan=3, pady=10)

        self.xRefreshButton = ttk.Button(bAutoRefreshButtonFrame, text="Refresh", command=self.refreshPARSNIP)
        self.xRefreshButton.grid(row=0, column=0, padx=5)

        self.bAutoRefreshButton = ttk.Button(bAutoRefreshButtonFrame, text="Enable Auto Refresh", command=self.toggleAutoRefreshPARSNIP)
        self.bAutoRefreshButton.grid(row=0, column=1, padx=5)

        # Frame for changes list
        self.xChangesFrame = ttk.Frame(self.root)
        self.xChangesFrame.grid(row=5, column=0, columnspan=3, sticky='nsew')

        # Treeview for changes list
        self.xChangesList = ttk.Treeview(self.xChangesFrame, columns=('Action', 'Description'), show='headings')
        self.xChangesList.heading('Action', text='Action')
        self.xChangesList.heading('Description', text='Description')

        self.xChangesList.column('Action', width=100, anchor='center')
        self.xChangesList.column('Description', width=800, anchor='w')

        vsbChanges = ttk.Scrollbar(self.xChangesFrame, orient="vertical", command=self.xChangesList.yview)
        self.xChangesList.configure(yscrollcommand=vsbChanges.set)

        self.xChangesList.grid(row=0, column=0, sticky='nsew')
        vsbChanges.grid(row=0, column=1, sticky='ns')

        # Label for loading status
        self.sLoadingLabel = ttk.Label(self.root, text="", anchor='center', font=('Arial', 10, 'italic'))
        self.sLoadingLabel.grid(row=6, column=0, columnspan=3, pady=10, sticky='s')

        # Grid configuration
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.xChangesFrame.grid_rowconfigure(0, weight=1)
        self.xChangesFrame.grid_columnconfigure(0, weight=1)

    def setHivePath(self):
        """Set the hive path based on user input."""
        sInputPath = self.sHivePathInputBox.get().strip()
        self.sHivePath = sInputPath
        messagebox.showinfo("Path Set", f"Hive path set to: {self.sHivePath}")

    def setEntryLimit(self):
        """Set the entry limit based on user input."""
        try:
            self.entry_limit = int(self.entryLimitInput.get().strip())
            messagebox.showinfo("Entry Limit Set", f"Entry limit set to: {self.entry_limit}")
        except ValueError:
            messagebox.showerror("Error", "Invalid entry limit. Please enter a valid number.")

    def setInterval(self):
        """Set the auto-refresh interval based on user input."""
        try:
            self.dInterval = int(self.intervalInput.get().strip())
            messagebox.showinfo("Interval Set", f"Auto-refresh interval set to: {self.dInterval} seconds")
        except ValueError:
            messagebox.showerror("Error", "Invalid interval. Please enter a valid number.")

    def exportRegistry(self):
        """Export unparsed Registry to the appropriate file in the script directory."""
        sHiveType = os.path.basename(self.sHivePath).split('.')[0].lower()
        sHiveExtension = os.path.splitext(self.sHivePath)[1]

        if sHiveType == 'ntuser':
            sHiveParameter = 'HKCU'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath))
        elif sHiveType == 'system':
            sHiveParameter = 'HKLM\\System'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + sHiveExtension)
        elif sHiveType == 'software':
            sHiveParameter = 'HKLM\\Software'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + sHiveExtension)
        elif sHiveType == 'sam':
            sHiveParameter = 'HKLM\\SAM'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + sHiveExtension)
        elif sHiveType == 'security':
            sHiveParameter = 'HKLM\\SECURITY'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + sHiveExtension)
        elif sHiveType == 'hardware':
            sHiveParameter = 'HKLM\\HARDWARE'
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + sHiveExtension)
        else:
            sHiveParameter = 'HKLM'  # Default to HKLM if not specified
            self.sExportPath = os.path.join(sScriptPath, os.path.basename(self.sHivePath) + sHiveExtension)

        try:
            subprocess.check_call(['reg', 'save', sHiveParameter, self.sExportPath, '/y'])
            messagebox.showinfo("Success", f"Unparsed Registry exported to: {self.sExportPath}")
        except subprocess.CalledProcessError as e:
            # If exporting fails due to access denied, try using VSS
            if 'access is denied' in str(e).lower():
                self.exportRegistryUsingVSS(sHiveType)
            else:
                messagebox.showerror("Error", f"Error exporting Parsed Registry: {e}")

    def exportRegistryUsingVSS(self, sHiveType):
        """Export Registry using VSS."""
        try:
            sShadowCopyPath = self.createVSSSnapshot()
            sShadowHivePath = os.path.join(sShadowCopyPath, f'Windows\\System32\\config\\{sHiveType}')
            sShadowExportPath = os.path.join(sScriptPath, f"{sHiveType}_shadow_copy.dat")

            subprocess.check_call(['reg', 'save', f'HKLM\\{sHiveType}', sShadowExportPath, '/y'])
            messagebox.showinfo("Success", f"Unparsed Registry exported from shadow copy to: {sShadowExportPath}")
            self.sExportPath = sShadowExportPath

        except Exception as e:
            messagebox.showerror("Error", f"Error exporting Parsed Registry using VSS: {e}")

    def createVSSSnapshot(self):
        """Create a VSS snapshot and return the shadow copy path."""
        try:
            subprocess.check_call(['vssadmin', 'create', 'shadow', '/for=C:'])
            xShadowList = subprocess.check_output(['vssadmin', 'list', 'shadows']).decode()
            for i in xShadowList.split('\n'):
                if "Shadow Copy Volume" in i:
                    sStart = i.find('{')
                    sEnd = i.find('}')
                    return f"\\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy{i[sStart+1:sEnd]}\\"
        except Exception as e:
            messagebox.showerror("Error", f"Error creating VSS snapshot: {e}")
            return None

    def parseRegistry(self, sHivePath):
        """Parse Registry using regipy."""
        xData = []
        try:
            with ThreadPoolExecutor() as executor:
                xHive = RegistryHive(sHivePath)
                for i in xHive.recurse_subkeys():
                    sKeyPath = i.path
                    for j in i.values:
                        xData.append({
                            "Key": sKeyPath,
                            "Name": j.name,
                            "Value": str(j.value),
                            "Type": j.value_type
                        })
        except Exception as e:
            messagebox.showerror("Error", f"Error parsing hive: {e}")
        return xData

    def loadGUITrees(self, xData):
        """Load parsed Registry data into Treeview."""
        for i in xData:
            parent = self.KeyTrees.insert('', 'end', text=i['Key'], open=False, tags=('key',))
            self.KeyTrees.insert(parent, 'end', values=(i['Name'], i['Value'], i['Type']), tags=('name', 'value', 'type'))

    def refreshPARSNIP(self):
        """Refresh the PARSNIP GUI manually."""
        self.sLoadingLabel.config(text="Loading...")
        self.root.update_idletasks()
        
        self.KeyTrees.delete(*self.KeyTrees.get_children())
        if self.isLiveHive(self.sHivePath):
            if self.sHivePath.lower() == sNtuserPath.lower():
                self.exportRegistry()
                sParsedPath = self.sExportPath  # Use the exported hive for parsing
            else:
                self.exportRegistry()
                sParsedPath = self.sExportPath  # Use the exported hive for parsing
        else:
            sParsedPath = self.sHivePath

        if os.path.exists(sParsedPath):
            sHiveType = os.path.basename(sParsedPath).split('.')[0].lower()
            if sHiveType != self.sPreviousHiveType:
                # Reset changes if hive to be parsed is different from the previous one
                self.xChangesList.delete(*self.xChangesList.get_children())
                self.xPreviousData = None  # Reset previous data to be none
                self.sPreviousHiveType = sHiveType
            
            data = self.parseRegistry(sParsedPath)
            self.loadGUITrees(data)
            
            # Compare changes only if there is a previous snapshot
            if self.xPreviousData:
                self.checkChanges(self.xPreviousData, data)
            
            self.xPreviousData = data
            self.exportToCSV(data, 'snapshot')

        self.sLoadingLabel.config(text="")

    def isLiveHive(self, hive_path):
        """Check if the hive path is in the system32 directory or is ntuser.dat."""
        sLowercaseHivePath = hive_path.lower()
        return 'system32' in sLowercaseHivePath or sLowercaseHivePath == sNtuserPath.lower()

    def toggleAutoRefreshPARSNIP(self):
        """Toggle auto-refresh functionality."""
        self.bAutoRefresh = not self.bAutoRefresh
        self.bAutoRefreshButton.config(text="Disable Auto Refresh" if self.bAutoRefresh else "Enable Auto Refresh")
        if self.bAutoRefresh:
            self.autoRefreshPARSNIP()

    def autoRefreshPARSNIP(self):
        """Auto-refresh the PARSNIP GUI at intervals."""
        if self.bAutoRefresh:
            self.refreshPARSNIP()
            self.root.after(self.dInterval * 1000, self.autoRefreshPARSNIP)

    def compareRegistrySnapshots(self, xOldData, xNewData):
        """Compare two snapshots of Registry data."""
        xChanges = []
        xOldDataSet = {f"{row['Key']}|{row['Name']}|{row['Value']}" for row in xOldData}
        xNewDataSet = {f"{row['Key']}|{row['Name']}|{row['Value']}" for row in xNewData}

        xAdded = xNewDataSet - xOldDataSet
        xRemoved = xOldDataSet - xNewDataSet

        for i in xAdded:
            xChanges.append({"Change": "Added", "Detail": i})

        for i in xRemoved:
            xChanges.append({"Change": "Removed", "Detail": i})

        return xChanges

    def checkChanges(self, xOldData, xNewData):
        """Check for changes in the Registry."""
        xChanges = self.compareRegistrySnapshots(xOldData, xNewData)

        self.xChangesList.delete(*self.xChangesList.get_children())

        for change in xChanges:
            self.xChangesList.insert('', 'end', values=(change['Change'], change['Detail']))
            
        xChangesData = [{"Action": self.xChangesList.set(item, 'Action'), "Description": self.xChangesList.set(item, 'Description')} for item in self.xChangesList.get_children()]
        self.exportToCSV(xChangesData, 'changes')

    def exportToCSV(self, xData, prefix):
        """Export parsed Registry data to CSV."""
        df = pd.DataFrame(xData)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_csv = os.path.join(sScriptPath, f"{prefix}_{timestamp}.csv")
        
        if len(df) > self.entry_limit:
            df = df.head(self.entry_limit)
        
        df.to_csv(output_csv, index=False)
        messagebox.showinfo("Export Complete", f"Data exported to: {output_csv}")

    def sortTreeview(self, sCol, bReverse):
        """Sort the Treeview using Python's sorted function by the given column."""
        xData = []
        
        if sCol == '#0':
            # Sort by Key column
            for item in self.KeyTrees.get_children(''):
                key = self.KeyTrees.item(item, 'text')
                xData.append((key, item))
            xDataSorted = sorted(xData, key=lambda item: item[0].lower(), reverse=bReverse)
        else:
            # Gather data based on column type for other columns
            for x in self.KeyTrees.get_children(''):
                key = self.KeyTrees.item(x, 'text')
                child_values = [self.KeyTrees.set(child, sCol) for child in self.KeyTrees.get_children(x)]
                sorted_child_value = sorted(child_values, reverse=bReverse)
                xData.append((sorted_child_value, key, x))

            # Sorting logic for non-key columns
            xDataSorted = sorted(xData, key=lambda item: item[0], reverse=bReverse)

        # Rearrange items in sorted positions
        for idx, data in enumerate(xDataSorted):
            self.KeyTrees.move(data[-1], '', idx)

        # Reverse sort next time
        self.KeyTrees.heading(sCol, command=lambda: self.sortTreeview(sCol, not bReverse))

        # Export sorted data
        self.exportSortedCSV(sCol)

    def exportSortedCSV(self, sCol):
        """Export sorted Treeview data to CSV with a name reflecting the sorting column."""
        sort_type = {'#0': 'KeySorted', 'Name': 'NameSorted', 'Value': 'ValueSorted', 'Type': 'TypeSorted'}
        xData = []
        for item in self.KeyTrees.get_children(''):
            key = self.KeyTrees.item(item, 'text')
            for child in self.KeyTrees.get_children(item):
                xData.append({
                    "Key": key,
                    "Name": self.KeyTrees.set(child, 'Name'),
                    "Value": self.KeyTrees.set(child, 'Value'),
                    "Type": self.KeyTrees.set(child, 'Type')
                })
        
        # Create a DataFrame and export
        df = pd.DataFrame(xData)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_csv = os.path.join(sScriptPath, f"snapshot_{sort_type[sCol]}_{timestamp}.csv")
        df.to_csv(output_csv, index=False)
        messagebox.showinfo("Export Complete", f"Sorted data exported to: {output_csv}")

# Main function to initialize and run the GUI
def main():
    root = tk.Tk()
    app = PARSNIP(root)
    root.mainloop()

if __name__ == "__main__":
    main()
