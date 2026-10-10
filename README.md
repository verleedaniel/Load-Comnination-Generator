# SAP2000 Load Combination Automation Using Excel VBA

## Overview

This tool automates the generation of load combinations in SAP2000 using Excel VBA and the SAP2000 API.

The macro reads load combination factors from an Excel worksheet and automatically creates the corresponding **linear additive response combinations** in SAP2000.

It also attempts to delete existing response combinations with matching generated names before recreating them with the updated factors.

The macro uses **late binding** to establish a connection with the SAP2000 API at runtime, eliminating the need to manually add SAP2000 API references in the VBA editor.

## Features

- **Excel-Based Input:** Reads load combination factors directly from an Excel worksheet.
- **Automated Generation:** Creates response combinations using the specified load factors.
- **Multiple Load Cases:** Supports Dead Load (DL), Imposed Load (IL), and earthquake load cases (Eqx, Eqy, Eqz).
- **Combination Management:** Deletes and recreates existing combinations with matching names.
- **One-Click Execution:** Runs the automation through a button assigned to the macro in Excel.
- **Late Binding:** Avoids manually adding SAP2000 API references in the VBA editor.

## Prerequisites

Before using the automation, ensure that you have:

- Microsoft Excel with VBA macro support.
- A compatible installation of SAP2000 with API access.
- The provided Excel macro-enabled workbook (`.xlsm`).
- A SAP2000 model containing the required load patterns and corresponding load cases.

## Setup and Usage

### Step 1: Create Load Patterns and Load Cases in SAP2000

Open your SAP2000 model and ensure that the required load patterns and corresponding load cases are defined.

Use the following exact names:

| Load Case Name | Description |
|---|---|
| `DL` | Dead Load |
| `IL` | Imposed Load |
| `Eqx` | Earthquake Load in the X-direction |
| `Eqy` | Earthquake Load in the Y-direction |
| `Eqz` | Earthquake Load in the Z-direction |

**Important:** The names must match exactly, including capitalization.

The VBA macro uses `SapModel.RespCombo.SetCaseList` to add load cases to each response combination. Therefore, `DL`, `IL`, `Eqx`, `Eqy`, and `Eqz` must exist as **load case names**, not merely as load pattern names.

Make sure each load case is properly defined and associated with the appropriate load pattern. You can create the load patterns and load cases through the SAP2000 interface before running the macro.

### Step 2: Download the Excel Workbook

1. Download the provided Excel macro-enabled workbook (`.xlsm`) from this GitHub repository.
2. Save the workbook to a trusted local folder on your computer.
3. Open the workbook using Microsoft Excel.

**Note:** Keep the file in `.xlsm` format to preserve the VBA macro and its assigned button.

### Step 3: Enable Macros

1. Open the workbook in Microsoft Excel.
2. If prompted, enable macros only if you trust the source of the workbook.
3. If Excel blocks the macros, check the file's security settings and follow your organization's policy for enabling trusted macros.

Do not enable macros in files from unknown or untrusted sources.

### Step 4: Enter Load Combination Factors

Enter the required load combination factors in the Excel worksheet using the following column arrangement:

| Column | Load Case | Description |
|---|---|---|
| A | DL | Dead Load factor |
| B | IL | Imposed Load factor |
| C | Eqx | Earthquake Load X factor |
| D | Eqy | Earthquake Load Y factor |
| E | Eqz | Earthquake Load Z factor |

**Input instructions:**

- Keep the column headers in Row 1.
- Enter the load combination factors starting from Row 2.
- Each subsequent row represents one load combination.
- Enter `0` for any load case that should not be included in a particular combination.
- Ensure all factors are numeric and correspond to the intended design code provisions.

### Step 5: Open the SAP2000 Model

Before running the macro:

1. Launch SAP2000.
2. Open the structural model you want to analyse.
3. Verify that the required load patterns and load cases are defined.
4. Save a backup of the model before executing the automation.

The macro attempts to connect to an existing SAP2000 instance. If it cannot connect, it attempts to start a new instance. Therefore, verify that the intended model is open and active before running the macro.

### Step 6: Generate Load Combinations

1. Return to the Excel workbook.
2. Click the button assigned to the VBA macro.
3. The macro connects to SAP2000 through its API.
4. It reads the load factors from each populated row of the worksheet.
5. It generates a combination name from the specified factors.
6. It attempts to delete any existing response combination with that name.
7. It creates a new linear additive response combination and adds the nonzero load case factors.

The process repeats for each populated row in the worksheet.

### Step 7: Verify the Results

After the macro finishes:

1. Open the response combinations section in SAP2000.
2. Verify that all expected combinations have been created.
3. Check the load cases included in each combination.
4. Confirm that the corresponding scale factors are correct.
5. Compare the generated combinations with the applicable design code requirements.
6. Save the SAP2000 model after verifying the results.

**Important:** The macro automates the creation of combinations based on the factors supplied in Excel. It does not independently determine or validate code-compliant load combinations.

## Important Notes

- **Load Case Names:** The load case names referenced in the VBA code must match those defined in SAP2000.
- **Load Patterns vs. Load Cases:** Defining a load pattern alone is insufficient. The corresponding load case must also exist.
- **Excel Format:** Use the `.xlsm` format to retain the macro.
- **Late Binding:** No manual SAP2000 API reference needs to be added in the VBA editor.
- **Existing Combinations:** Combinations with matching generated names may be deleted and recreated. Back up your model before execution.
- **Input Validation:** Check that all load factors are numeric and that the worksheet contains valid input data.
- **Design Code Compliance:** Verify the load factors, signs, and combinations against the applicable design code before using them for structural design.

## Technologies Used

- Microsoft Excel
- Visual Basic for Applications (VBA)
- SAP2000 API
- COM Automation
- Late Binding

## Disclaimer

This tool is intended to simplify repetitive tasks in structural engineering workflows by automating a specific part of the load combination generation process.

The user is responsible for verifying the correctness of the load factors, load cases, generated combinations, and applicable design code requirements before relying on the results for engineering analysis or design.

Always maintain a backup of the original SAP2000 model before running the automation.

## Authors

**Verlee Daniel Gudapati**  
**Janesh Naredla**
