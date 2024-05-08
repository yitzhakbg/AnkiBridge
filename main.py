import tkinter as tk
from tkinter import filedialog, messagebox, Label, Button, Frame
import pandas as pd
import pandas as pd

def parse_excel_file(excel_path):
    # Read the Excel file
    df = pd.read_excel(excel_path)
    
    # Assume the Excel file has columns 'First Name' and 'Family Name'
    if 'First Name' not in df.columns or 'Family Name' not in df.columns:
        raise ValueError("Excel file must contain 'First Name' and 'Family Name' columns.")
    
    # Create profile names
    df['Profile Name'] = df['First Name'].str.lower() + '.' + df['Family Name'].str.lower().str[0]
    
    return df[['First Name', 'Family Name', 'Profile Name']]

# This function will be integrated into the GUI application to process the selected Excel file.


class AnkiProfileCreatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Anki Profile Creator")
        
        self.frame = Frame(self.root)
        self.frame.pack(padx=10, pady=10)

        self.excel_label = Label(self.frame, text="Select Excel File:")
        self.excel_label.grid(row=0, column=0, sticky='w')
        
        self.excel_button = Button(self.frame, text="Browse", command=self.load_excel)
        self.excel_button.grid(row=0, column=1, padx=5, pady=5)

        self.anki_label = Label(self.frame, text="Select Anki Collection File:")
        self.anki_label.grid(row=1, column=0, sticky='w')
        
        self.anki_button = Button(self.frame, text="Browse", command=self.load_anki)
        self.anki_button.grid(row=1, column=1, padx=5, pady=5)

        self.generate_button = Button(self.frame, text="Generate Profiles", command=self.generate_profiles)
        self.generate_button.grid(row=2, columnspan=2, pady=10)

        self.excel_path = None
        self.anki_path = None

    def load_excel(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.excel_path:
            self.excel_label.config(text=f"Selected: {os.path.basename(self.excel_path)}")

    def load_anki(self):
        self.anki_path = filedialog.askopenfilename(filetypes=[("Anki Collection", "collection.anki2")])
        if self.anki_path:
            self.anki_label.config(text=f"Selected: {os.path.basename(self.anki_path)}")

def generate_profiles(self):
        if not self.excel_path or not self.anki_path:
            messagebox.showerror("Error", "Please select both an Excel file and an Anki collection file.")
            return
        
        try:
            # Integration of components with enhanced error handling
            profiles_df = parse_excel_file(self.excel_path)
            profile_names = profiles_df['Profile Name'].tolist()
            profiles_to_create = check_existing_profiles(self.anki_path, profile_names)
            create_new_profiles(profiles_to_create, self.anki_path)
            
            messagebox.showinfo("Success", "Profiles generated successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    self.generate_button.config(command=generate_profiles)

# This function will be called to integrate all components in the GUI application.


if __name__ == "__main__":
    root = tk.Tk()
    app = AnkiProfileCreatorApp(root)
    root.mainloop()

from aqt.profiles import ProfileManager

def check_existing_profiles(anki_path, profile_names):
    profile_manager = ProfileManager()
    profile_manager.load(anki_path)
    existing_profiles = set(profile_manager.all_profiles())

    # Determine which profiles need to be created
    profiles_to_create = [name for name in profile_names if name not in existing_profiles]
    return profiles_to_create
from aqt import mw
from aqt.profiles import ProfileManager

def create_new_profiles(profiles_to_create, anki_path):
    profile_manager = ProfileManager()
    profile_manager.load(anki_path)
    
    for profile_name in profiles_to_create:
        if profile_name not in profile_manager.all_profiles():
            profile_manager.create(profile_name)
            print(f"Profile created: {profile_name}")
        else:
            print(f"Profile already exists: {profile_name}")

    profile_manager.save()

# This function will be integrated into the GUI application to create new profiles as needed.




