import os
import shutil

# Define multiple source network paths (list of directories)
source_network_paths = [
    r'\\192.168.1.14\Essential Software\6-SAP\SAP-AppData files\Common',
    r'\\192.168.1.14\Essential Software\6-SAP\SAP-AppData files\LogonServerConfigCache',
    r'\\192.168.1.14\Essential Software\6-SAP\SAP-AppData files\SAP GUI'
]

# Get the AppData folder for the current user dynamically
appdata_folder = os.getenv('APPDATA')
destination_local_path = os.path.join(appdata_folder, 'SAP')  # Local destination folder in AppData

# Ensure destination directory exists
os.makedirs(destination_local_path, exist_ok=True)

# Function to copy a file (directly replaces if exists)
def copy_file(source_file, destination_file):
    try:
        shutil.copy(source_file, destination_file)  # Copy file, replaces if exists
        print(f"Successfully copied file: {source_file} to {destination_file}")
    except Exception as e:
        print(f"Error copying file {source_file}: {e}")

# Function to remove a directory and its contents (forcefully if needed)
def remove_directory(directory):
    try:
        # Attempt to remove the directory (force remove files first)
        for root, dirs, files in os.walk(directory, topdown=False):
            for name in files:
                file_path = os.path.join(root, name)
                try:
                    os.remove(file_path)  # Remove file
                    print(f"Removed file: {file_path}")
                except PermissionError:
                    print(f"Permission denied for file: {file_path}")
                except Exception as e:
                    print(f"Error removing file {file_path}: {e}")

            for name in dirs:
                dir_path = os.path.join(root, name)
                try:
                    os.rmdir(dir_path)  # Remove directory
                    print(f"Removed directory: {dir_path}")
                except Exception as e:
                    print(f"Error removing directory {dir_path}: {e}")

        # Now remove the main directory itself
        os.rmdir(directory)
        print(f"Successfully removed directory: {directory}")

    except Exception as e:
        print(f"Error removing directory {directory}: {e}")

# Function to copy a directory and its contents (removes existing directory before copying)
def copy_directory(source_dir, destination_dir):
    if os.path.exists(destination_dir):
        try:
            # Remove existing directory before copying
            remove_directory(destination_dir)
        except Exception as e:
            print(f"Error removing existing directory {destination_dir}: {e}")
            return

    try:
        shutil.copytree(source_dir, destination_dir)  # Copies the entire directory
        print(f"Successfully copied directory: {source_dir} to {destination_dir}")
    except Exception as e:
        print(f"Error copying directory {source_dir}: {e}")

# Function to copy either file or directory
def copy_from_network(source_path, destination_path):
    # Check if source path exists
    if not os.path.exists(source_path):
        print(f"Error: The source path does not exist: {source_path}")
        return

    # Check if the source is a file or directory
    if os.path.isfile(source_path):
        # If it's a file, copy it
        destination_file = os.path.join(destination_path, os.path.basename(source_path))
        copy_file(source_path, destination_file)

    elif os.path.isdir(source_path):
        # If it's a directory, copy it
        destination_dir = os.path.join(destination_path, os.path.basename(source_path))
        copy_directory(source_path, destination_dir)

    else:
        print(f"Error: The source path is neither a file nor a directory: {source_path}")

# Loop through each source path and copy it
for source_path in source_network_paths:
    copy_from_network(source_path, destination_local_path)
