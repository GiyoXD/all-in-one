import os
import subprocess
import argparse
import sys
from pathlib import Path
import logging
from typing import List  # Add this import for proper type hints
import shutil

# Setup basic logging for the wrapper script
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def run_script(script_path: Path, args: List[str] = [], cwd: Path = None, script_name: str = "") -> bool:
    """Runs a Python script using subprocess, handling potential errors."""
    if not script_path.is_file():
        logging.error(f"Script not found: {script_path}")
        return False

    command = [sys.executable, str(script_path)] + args
    script_name = script_name or script_path.name
    logging.info(f"Running {script_name}...")
    logging.info(f"  Command: {' '.join(command)}")
    if cwd:
        # Ensure CWD exists before running
        if not Path(cwd).is_dir():
             logging.error(f"Working directory not found for {script_name}: {cwd}")
             return False
        logging.info(f"  Working Directory: {cwd}")

    try:
        # Use encoding and error handling for robustness
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=True,
            cwd=cwd,
            encoding='utf-8',
            errors='replace'
        )

        stdout_content = result.stdout.strip() if result.stdout else ""
        stderr_content = result.stderr.strip() if result.stderr else ""

        if stdout_content:
            logging.info(f"{script_name} Output:\n--- START ---\n{stdout_content}\n--- END ---")
        else:
             logging.info(f"{script_name} produced no standard output.")

        if stderr_content:
            # Treat stderr as warning unless script explicitly failed (check=True handles that)
            logging.warning(f"{script_name} Stderr:\n--- START ---\n{stderr_content}\n--- END ---")

        logging.info(f"{script_name} completed successfully.")
        return True

    except subprocess.CalledProcessError as e:
        logging.error(f"Error running {script_name} (Return Code: {e.returncode}):")
        stdout_content = e.stdout.strip() if e.stdout else "No stdout captured."
        stderr_content = e.stderr.strip() if e.stderr else "No stderr captured."
        logging.error(f"Stdout:\n{stdout_content}")
        logging.error(f"Stderr:\n{stderr_content}")
        return False
    except FileNotFoundError:
        # This might catch python executable not found, unlikely for script path due to initial check
        logging.error(f"Error: Could not find executable '{sys.executable}' or script '{script_path}' during execution.")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred while trying to run {script_name}: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Automate JSON creation and Invoice generation from an input Excel file."
    )
    parser.add_argument(
        "-i", "--input",
        required=True,
        help="Path to the input Excel file (e.g., 'input.xlsx').",
        type=str
    )
    parser.add_argument(
        "--fob",
        action="store_true",
        help="Pass --fob flag to invoice_gen/generate_invoice.py."
    )
    parser.add_argument(
        "--custom",
        action="store_true",
        help="Pass --custom flag to invoice_gen/generate_invoice.py."
    )

    args = parser.parse_args()
    
    # Get the input file path and create output directory based on input filename
    input_excel_path = Path(args.input).resolve()
    if not input_excel_path.is_file():
        logging.error(f"Input Excel file not found: {input_excel_path}")
        sys.exit(1)
    
    # Create output directory with same name as input file (without extension)
    output_dir = input_excel_path.parent / input_excel_path.stem
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Define all required directories relative to the output directory
    json_output_dir = output_dir / "json_output"
    invoice_output_dir = output_dir / "invoice_output"
    
    # Create required directories
    json_output_dir.mkdir(parents=True, exist_ok=True)
    invoice_output_dir.mkdir(parents=True, exist_ok=True)
    
    # --- Define Project Structure & Validate Paths ---
    project_root = Path(__file__).parent.resolve()
    create_json_dir = project_root / "create_json"
    invoice_gen_dir = project_root / "invoice_gen"
    create_json_script = create_json_dir / "main.py"
    invoice_gen_script = invoice_gen_dir / "generate_invoice.py"
    template_dir = project_root / "invoice_gen" / "TEMPLATE"
    config_dir = project_root / "invoice_gen" / "config"
    
    # Get the identifier from the input filename
    identifier = input_excel_path.stem
    
    # Extract the prefix (e.g., "JF" from "JF25001")
    prefix = ''.join(filter(str.isalpha, identifier))
    if not prefix:
        logging.error(f"Could not extract prefix from filename: {identifier}")
        sys.exit(1)

    # Validate paths that must exist beforehand
    if not create_json_script.is_file():
        logging.error(f"JSON creation script not found: {create_json_script}")
        sys.exit(1)
    if not invoice_gen_script.is_file():
        logging.error(f"Invoice generation script not found: {invoice_gen_script}")
        sys.exit(1)
    if not template_dir.is_dir():
        logging.error(f"Template directory not found: {template_dir}")
        sys.exit(1)
    if not config_dir.is_dir():
        logging.error(f"Configuration directory not found: {config_dir}")
        logging.error("Please ensure the directory exists and is correct.")
        sys.exit(1)

    # --- Step 1: Run create_json/main.py ---
    create_json_args = [
        "--input-excel", str(input_excel_path),
        "--output-dir", str(json_output_dir)
    ]
    logging.info(f"Running JSON creation step using input: {input_excel_path}")
    if not run_script(create_json_script, args=create_json_args, cwd=create_json_dir, script_name="create_json"):
        logging.error("JSON creation script failed. Aborting.")
        sys.exit(1)

    # --- Step 2: Verify JSON Output ---
    expected_json_path = json_output_dir / f"{identifier}.json"
    if not expected_json_path.is_file():
        logging.error(f"Expected JSON output file was not found: {expected_json_path}")
        logging.error("Check the output/logs of the create_json script for errors.")
        sys.exit(1)
    logging.info(f"JSON file found: {expected_json_path}")

    # --- Step 3: Verify Expected Config for invoice_gen ---
    expected_config_path = config_dir / f"{prefix}_config.json"
    logging.info(f"Invoice generation step will expect config file: {expected_config_path}")
    if not expected_config_path.is_file():
        logging.error(f"Expected config file '{expected_config_path}' not found in '{config_dir}'.")
        logging.error("Please ensure the required config file exists.")
        sys.exit(1)

    # --- Step 4: Run invoice_gen/generate_invoice.py for each mode ---
    modes = [
        ("normal", []),
        ("fob", ["--fob"]),
        ("custom", ["--custom"])
    ]
    
    for mode, flags in modes:
        logging.info(f"--- Processing {mode.upper()} mode ---")
        
        # Prepare invoice generation arguments with new naming convention
        output_filename = f"CT&INV&PL {identifier} {mode.upper()}.xlsx"
        invoice_gen_args = [
            str(expected_json_path),
            "--output", str(invoice_output_dir / output_filename),
            "--templatedir", str(template_dir),
            "--configdir", str(config_dir),
        ] + flags
        
        logging.info(f"Running Invoice generation step to create: {output_filename}")
        if not run_script(invoice_gen_script, args=invoice_gen_args, cwd=invoice_gen_dir, script_name="invoice_gen"):
            logging.error(f"Invoice generation script failed for {mode} mode.")
            continue

    logging.info("--- Automation Completed Successfully ---")
    logging.info(f"All outputs saved in directory: {output_dir}")
    logging.info("Generated three versions:")
    logging.info(f"1. Normal: CT&INV&PL {identifier} NORMAL.xlsx")
    logging.info(f"2. FOB: CT&INV&PL {identifier} FOB.xlsx")
    logging.info(f"3. Custom: CT&INV&PL {identifier} CUSTOM.xlsx")

if __name__ == "__main__":
    main()
