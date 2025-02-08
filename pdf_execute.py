import pandas as pd
from excel_parser import eligible_service_members, ineligible_service_members
from pdf_generator import generate_roster_pdf
from datetime import datetime


def execute_roster_generation(alpha_roster_path, cycle, year, output_path=None):
    """
    Execute the roster generation process.

    Args:
        alpha_roster_path (str): Path to the Alpha Roster Excel file
        cycle (str): Promotion cycle (e.g., 'SSG')
        year (int): Year for promotion cycle
        output_path (str, optional): Path for output PDF. If None, generates default name
    """
    try:
        # Define required columns
        required_columns = [
            'FULL_NAME', 'GRADE', 'ASSIGNED_PAS_CLEARTEXT', 'DAFSC', 'DOR',
            'DATE_ARRIVED_STATION', 'TAFMSD', 'REENL_ELIG_STATUS', 'ASSIGNED_PAS'
        ]
        optional_columns = ['GRADE_PERM_PROJ', 'UIF_CODE', 'UIF_DISPOSITION_DATE']
        pdf_columns = [
            'FULL_NAME', 'GRADE', 'DATE_ARRIVED_STATION', 'DAFSC',
            'ASSIGNED_PAS_CLEARTEXT', 'DOR', 'TAFMSD', 'ASSIGNED_PAS'
        ]

        # Read Excel file
        print(f"Reading Alpha Roster from: {alpha_roster_path}")
        alpha_roster = pd.read_excel(alpha_roster_path)

        # Filter columns
        filtered_alpha_roster = alpha_roster[required_columns + optional_columns]
        pdf_roster = filtered_alpha_roster[pdf_columns]

        # Create eligible and ineligible dataframes
        eligible_df = pdf_roster.loc[eligible_service_members]
        ineligible_df = pdf_roster.loc[ineligible_service_members]

        # Generate output filename if not provided
        if output_path is None:
            current_date = datetime.now().strftime("%Y%m%d")
            output_path = f"promotion_roster_{cycle}_{current_date}.pdf"

        # Generate PDF
        print(f"Generating PDF: {output_path}")
        print(f"Eligible members: {len(eligible_df)}")
        print(f"Ineligible members: {len(ineligible_df)}")

        generate_roster_pdf(
            eligible_df=eligible_df,
            ineligible_df=ineligible_df,
            cycle=cycle,
            output_filename=output_path
        )

        print(f"PDF generation complete: {output_path}")
        return True

    except Exception as e:
        print(f"Error generating roster: {str(e)}")
        return False


if __name__ == "__main__":
    # Configuration
    ALPHA_ROSTER_PATH = r'C:\Users\Gamer\Downloads\Alpha Roster.xlsx'
    CYCLE = 'SSG'
    YEAR = 2025

    # Execute roster generation
    success = execute_roster_generation(
        alpha_roster_path=ALPHA_ROSTER_PATH,
        cycle=CYCLE,
        year=YEAR
    )

    if success:
        print("Roster generation completed successfully.")
    else:
        print("Roster generation failed. Check error messages above.")