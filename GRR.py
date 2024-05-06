import os
import pandas as pd
import numpy as np

# Specify the parent folder path where the subfolders are located
parent_folder_path = os.path.dirname(os.path.abspath(__file__))

# Get a list of all subfolders in the parent folder
subfolders = [folder for folder in os.listdir(parent_folder_path) if os.path.isdir(os.path.join(parent_folder_path, folder))]

# Initialize an empty list to store the 3D matrices
matrices = []

# Iterate over each subfolder
for subfolder in subfolders:
    # Construct the subfolder path
    subfolder_path = os.path.join(parent_folder_path, subfolder)
    
    # Get a list of all Excel files in the subfolder
    excel_files = [file for file in os.listdir(subfolder_path) if file.endswith('.xlsx')]

    # Initialize an empty list to store the data frames
    data_frames = []

    # Iterate over each Excel file
    for file in excel_files:
        # Construct the file path
        file_path = os.path.join(subfolder_path, file)
        
        # Read the Excel file into a data frame
        df = pd.read_excel(file_path, header=None)
        df = df.values.flatten()
        
        # Append the data frame to the list
        data_frames.append(df)

    # Convert the list of data frames into a 3D numpy array
    combined_matrix = np.array([df for df in data_frames])
    # Append the matrix to the list
    matrices.append(combined_matrix)

combined_matrix = np.array(matrices)

# Calculate the mean of each part for each operator
operator_means = np.mean(combined_matrix, axis=1)

# Calculate the mean of each operator for each part
part_means = np.mean(combined_matrix, axis=0)

# Calculate the total mean
total_mean = np.mean(combined_matrix)

# Calculate the repeatability
repeatability = np.mean((combined_matrix - operator_means[:, np.newaxis, :])**2)

# Calculate the reproducibility
reproducibility = np.mean((operator_means - total_mean)**2)

# Calculate the part-to-part variance
part_to_part_variance = np.mean((part_means - total_mean)**2)

# Calculate the total GRR
grr = repeatability + reproducibility

num_operators = combined_matrix.shape[0]
print(f"Number of operators: {num_operators}")
num_trials = combined_matrix.shape[1]
print(f"Number of trials: {num_trials}")
num_parts = combined_matrix.shape[2]
print(f"Number of parts: {num_parts}")
print("\n")
print(f"Repeatability: {repeatability}")
print(f"Reproducibility: {reproducibility}")
print(f"Part-to-part variance: {part_to_part_variance}")
print(f"GRR: {grr}")


