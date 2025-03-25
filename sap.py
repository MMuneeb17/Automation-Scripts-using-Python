import os
import shutil

# Define source and destination paths
source_network_path = r'\\192.168.1.14\Essential Software'  # Network location
destination_local_path = r'C:\Users\IT\Desktop'  # Local destination folder

# Ensure destination directory exists
os.makedirs(destination_local_path, exist_ok=True)

# Check if source path is valid
if not os.path.isdir(source_network_path):
    print(f"Error: The source path is not a valid directory: {source_network_path}")
else:
    # List all files in the source directory
    try:
        files_to_copy = os.listdir(source_network_path)
        print(f"Found {len(files_to_copy)} items in the source directory.")

        # Copy each file from the network to the local destination
        for file_name in files_to_copy:
            # Create full source and destination paths
            source_file = os.path.join(source_network_path, file_name)
            destination_file = os.path.join(destination_local_path, file_name)

            # Debugging: Print the full source and destination paths
            print(f"Source: {source_file}")
            print(f"Destination: {destination_file}")

            # Check if it's a file (not a directory) and then copy
            if os.path.isfile(source_file):
                try:
                    shutil.copy(source_file, destination_file)
                    print(f"Successfully copied: {file_name}")
                except Exception as e:
                    print(f"Error copying {file_name}: {e}")
            else:
                print(f"Skipping directory: {file_name}")

    except Exception as e:
        print(f"Error accessing source directory: {e}")
