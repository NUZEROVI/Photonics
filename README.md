# Data Analysis & Create Apps (For Paul)

## (1) Electrical Conductance (EC) dataset analysis
- Step 1 - EC - App1 

<p align="center">
    <img src="https://github.com/analyzeDataVis/Photonics/blob/1_Electrical_Conductance/1_Electrical_Conductance/Final%20Task%20(Combine%20ALL%20Tasks)/Sceenshot/App1.png">
</p>


Input Files : [multiple .CSV Files (MFCV-1, MFCV-2, MFCV-3, MFCV-4)](https://github.com/analyzeDataVis/Photonics/tree/1_Electrical_Conductance/1_Electrical_Conductance/Requirements/Final%20Datasets%20(ALL%20Tasks))
-> Auto grab VBias

Output Files :        

        1. CV and GV Sheets         
        2. Gp/ω Sheet (with noise data)
    
- Step 2 - Noise by manual inspection : removal noise rows and columns from Gp/ω sheet

- Step 3 - EC - App2

<p align="center">
    <img src="https://github.com/analyzeDataVis/Photonics/blob/1_Electrical_Conductance/1_Electrical_Conductance/Final%20Task%20(Combine%20ALL%20Tasks)/Sceenshot/App2_v3%20(readme).png">
</p>

Input Files : Gp/ω sheet with removal noise dataset (.xlsx)

Output Files :  

        poly peek X-Y coordinate  (.xlsx)

### [Description with how to find peak and enhance accuracy.](https://github.com/analyzeDataVis/Photonics/blob/1_Electrical_Conductance/1_Electrical_Conductance/Task2%20(Curve%20Fitting%20%2B%20Find%20Peak)/Analysis%20(task2).ipynb)
<p align="center">
    <img src="https://github.com/analyzeDataVis/Photonics/blob/Interactive-Charts-with-D3/1_Electrical_Conductance/Task2%20(Curve%20Fitting%20%2B%20Find%20Peak)/Sceenshots/description.png">
</p>

### [Additional Description](https://github.com/analyzeDataVis/Photonics/tree/main/1_Electrical_Conductance/Interactive_D3)
if the dataset has the tolerance, we can find the tolerance information by interactive d3 tool.
<p align="center">
    <img src="https://github.com/analyzeDataVis/Photonics/blob/Interactive-Charts-with-D3/1_Electrical_Conductance/Interactive_D3/Demo/d3_scatter_interactive.gif">
</p>
