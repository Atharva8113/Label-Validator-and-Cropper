# Label Validator & Cropper (FitLabel)
Automates the validation of shipping labels against invoices and ADC documents, checking for correct Lot Numbers and Import Licenses before cropping and renaming them for shipment.

## Quick Start
*   **Requirements:** Windows OS (No Python installation needed).
*   **Installation:** Download the folder and locate the `FitLabel_GUI.exe` file.

## How to Use (Step-by-Step)
1.  **Launch the App**: Double-click `FitLabel_GUI.exe` to open the Label Validator & Cropper.
2.  **Load Invoices**: Click the **Select** button under "📄 Invoice PDFs" and choose your invoice file(s). The app will extract Item and Lot numbers.
3.  **Load Labels**: Click the **Select** button under "🏷️ Label PDFs" to upload the labels you need to process. The app automatically checks if they match the invoice data.
4.  **Verify with ADC (Optional)**: Click **Select** under "📋 ADC PDF" to load current ADC data. This adds a verification step for the Import License number.
5.  **Review Preview**: Check the "Preview" pane.
    *   **Green (✓)**: Valid matches ready for processing.
    *   **Red (✗)**: Mismatches or missing data.
6.  **Generate**: Click the **📥 Generate Cropped Labels** button.
7.  **Save**: A dialog will appear asking where to save the files. Select a destination folder. The tool will crop the labels, rename them to `[List No] [Lot No].pdf`, and save them there.

## Common Issues
*   **Windows Defender Warning**: If Windows Defender prevents the app from starting, click **'More Info' -> 'Run Anyway'**. This happens because it is an internal tool.
*   **Text Recognition**: Ensure your PDFs are text-readable (not scanned images). If you see "No item-lot pairs found", the tool might not be able to read the text in your Invoice/Label PDF.

## Contact
For support or feature requests, please contact the IT Team.
