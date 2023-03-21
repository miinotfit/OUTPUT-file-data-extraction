import os
import pandas as pd

# Define the directory containing the .out files
directory = r"C:\Users\Kenneth\Desktop\code test"

# Initialize empty lists to hold the data
sa_sf_rpa_data = []
sf_dft_data = []

# Loop over all files in the directory
for filename in os.listdir(directory):
    if filename.endswith(".out"):
        # Open the file
        with open(os.path.join(directory, filename), 'r') as f:
            # Read in the lines of the file
            lines = f.readlines()
            # Find the pivot point
            pivot = None
            for i, line in enumerate(lines):
                if "OPTIMIZATION CONVERGED" in line:
                    pivot = i
                    break
            # Find the SA-SF-RPA Excitation Energies after the pivot
            if pivot is not None:
                for i in range(pivot, len(lines)):
                    state_numRPA = 0
                    if "SA-SF-RPA Excitation Energies" in lines[i]:
                        print("SA-SF-RPA Excitation Energies found in file:", filename)
                        # Extract the numerical value after eV for each excited state
                        for j in range(i+1, len(lines)):
                            if "Total energy for state" in lines[j]:
                                state_numRPA += 1 # increment excited state number
                                sa_sf_rpa_val = float(lines[j].split(':')[1].strip())
                                sa_sf_rpa_data.append((filename, state_numRPA, sa_sf_rpa_val))
                            if "Excited state   4:" in lines[j]:  # Stop after the 4th excited state
                                break
            else:
                print("Optimization Converged not found in file:", filename)
            # Find the SF-DFT Excitation Energies before the pivot
            if pivot is not None:
                for i in range(pivot, -1, -1):
                    state_numDFT = 0 # initialize excited state number
                    if "SF-DFT Excitation Energies" in lines[i]:
                        print("SF-DFT Excitation Energies found in file:", filename)
                        # Extract the numerical value after <S**2> for each excited state
                        for j in range(i, len(lines)):
                            if "Excited state" in lines[j]:
                                # Extract the number after <S**2>
                                for line in lines[j:]:
                                    if "<S**2>" in line:
                                        state_numDFT += 1 # increment excited state number
                                        s_squared = float(line.split()[2])
                                        sf_dft_data.append((filename, state_numDFT, s_squared))
                            elif "Total energy for state" in lines[j]:
                                # End of the excited states section
                                break
                        break
            else:
                print("Optimization Converged not found in file:", filename)

# Create pandas dataframes from the lists
sa_sf_rpa_df = pd.DataFrame(sa_sf_rpa_data, columns=["filename", "excited state number", "eV"])
sf_dft_df = pd.DataFrame(sf_dft_data, columns=["filename", "excited state number", "<S**2>"])

# Write the dataframes to an Excel file
with pd.ExcelWriter('output.xlsx') as writer:
    sa_sf_rpa_df.to_excel(writer, sheet_name='SA-SF-RPA Excitation Energies', index=False)
    sf_dft_df.to_excel(writer, sheet_name='SF-DFT Excitation Energies', index=False)