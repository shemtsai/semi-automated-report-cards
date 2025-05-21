# Semi-automated Stewardship Report Cards

The R code processes antibiotic administration data from an xlsx file, transforms the data, and generates personalized report cards as Word documents. The code can be modified as needed per ASP/institution

## Files

- **Automated reports.R**  
  Main R script that creates pivot tables, de-identifies provider names, and generates reports for each provider.

- **Data Structure Example.xlsx**  
  Serves as a sample template for the input data as an example of what structure should look like for code to run. Data is de-identified and randomized. Physician names and Patient IDs are fictional; any resemblance or similarity to actual persons, living or dead, is purely coincidental. 

## Requirements

- **R Packages**  
- The following packages are needed. The script automatically installs the required R packages if they are not already installed:
  - `readxl`
  - `MASS`
  - `ggplot2`
  - `officer`
  - `flextable`
  - `tidyr`
  - `dplyr`
  - `magrittr`
  - `writexl`

## Usage

1. **Prepare your Data**  
   Use the provided "Data Structure Example.xlsx" as a guide for formatting your input Excel file with antibiotic administration data.

2. **Modify the Script**  
    Adjust the script as needed for the OrderingPhysicianNames and Routes to match your institution-specific data. Further sections that may be beneficial to modify are listed with a #NOTE in the code but any part of the code can be customized including metrics of interest. 

3. **Run the Script**  
   Open the script file in R and execute it. The script will prompt you to:
   - Select your input Excel file.
   - Choose a directory where the personalized reports will be saved.

4. **Review Output**  
   The script generates several pivot tables, plots, and creates de-identified, personalized report cards.

## Contact

For questions, suggestions, or collaborations, please contact stsai@houstonmethodist.org
