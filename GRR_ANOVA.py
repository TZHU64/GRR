import os
import pandas as pd
import numpy as np
import statsmodels.api as sm
from statsmodels.formula.api import ols

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

num_operators, num_trials, num_parts = combined_matrix.shape

# Create a DataFrame from the 3D matrix
data = {
    'Operator': np.repeat(np.arange(num_operators), num_trials * num_parts),
    'Trial': np.tile(np.repeat(np.arange(num_trials), num_parts), num_operators),
    'Part': np.tile(np.arange(num_parts), num_operators * num_trials),
    'Measurement': combined_matrix.flatten(),
}
df = pd.DataFrame(data)

# Perform an ANOVA
model = ols('Measurement ~ C(Operator) + C(Part)', df).fit()
anova_table = sm.stats.anova_lm(model, typ=2)

# Calculate the repeatability, reproducibility, and part-to-part variance
repeatability = model.scale
reproducibility = anova_table.loc['C(Operator)', 'sum_sq'] / (num_operators - 1)
part_to_part_variance = anova_table.loc['C(Part)', 'sum_sq'] / (num_parts - 1)

# Calculate the total GRR
grr = repeatability + reproducibility

print(f"Number of operators: {num_operators}")
print(f"Number of trials: {num_trials}")
print(f"Number of parts: {num_parts}")
print("\nANOVA Table:")
print(anova_table)
print("\nGRR Components:")
print(f"Repeatability: {repeatability}")
print(f"Reproducibility: {reproducibility}")
print(f"Part-to-part variance: {part_to_part_variance}")
print(f"GRR: {grr}")