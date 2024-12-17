import os
import shutil

def delete_folder(folder_path):
    try:
        # Check if the folder exists
        if not os.path.exists(folder_path):
            return f"Folder '{folder_path}' does not exist."
        
        # Delete the folder and its contents
        shutil.rmtree(folder_path)
        print(f"Folder '{folder_path}' and its contents have been deleted.")
    except FileNotFoundError:
        print("The folder does not exist.")
    except Exception as e:
        return f"Error: {e}"
    