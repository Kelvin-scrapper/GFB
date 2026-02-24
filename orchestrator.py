#!/usr/bin/env python3
"""
GFB Data Processing Orchestrator

This script orchestrates the execution of the GFB data extraction pipeline,
providing a centralized entry point for processing German Federal Budget data.
"""

import os
import sys
import subprocess
import argparse
from datetime import datetime
from pathlib import Path

def print_banner():
    """Print application banner"""
    print("=" * 60)
    print("         GFB Data Processing Orchestrator")
    print("    German Federal Budget Data Extraction Pipeline")
    print("=" * 60)
    print()

def check_dependencies():
    """Check if required dependencies are installed"""
    required_packages = ['pandas', 'numpy', 'openpyxl', 'selenium', 'undetected_chromedriver', 'requests']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"ERROR: Missing required packages: {', '.join(missing_packages)}")
        print("Please install them using: pip install -r requirements.txt")
        return False
    
    return True

def check_files():
    """Check if required script files exist"""
    required_files = ['main.py', 'map.py']
    missing_files = []
    
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print(f"ERROR: Missing required files: {', '.join(missing_files)}")
        return False
    
    return True

def run_download():
    """Run the universal download script (main.py)"""
    print("Running universal download script (main.py)...")
    print("-" * 40)
    
    try:
        # Set environment variable to handle Unicode properly
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        
        # Run main.py and capture output
        result = subprocess.run([sys.executable, 'main.py'], 
                              capture_output=True, 
                              text=True, 
                              check=True,
                              env=env,
                              encoding='utf-8',
                              errors='replace')
        
        # Print the output (replace problematic Unicode chars)
        if result.stdout:
            stdout_clean = result.stdout.replace('✅', 'SUCCESS').replace('❌', 'ERROR').replace('📁', '').replace('📏', '')
            print(stdout_clean)
        
        if result.stderr:
            stderr_clean = result.stderr.replace('✅', 'SUCCESS').replace('❌', 'ERROR').replace('📁', '').replace('📏', '')
            print("STDERR:", stderr_clean)
        
        print("-" * 40)
        print("Download completed successfully!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Download failed with return code {e.returncode}")
        if e.stdout:
            stdout_clean = e.stdout.replace('✅', 'SUCCESS').replace('❌', 'ERROR').replace('📁', '').replace('📏', '')
            print("STDOUT:", stdout_clean)
        if e.stderr:
            stderr_clean = e.stderr.replace('✅', 'SUCCESS').replace('❌', 'ERROR').replace('📁', '').replace('📏', '')
            print("STDERR:", stderr_clean)
        return False
    except Exception as e:
        print(f"ERROR: Unexpected error during download: {str(e)}")
        return False

def run_custom_download(url, keywords):
    """Run download from custom URL with keywords"""
    print(f"Running custom download from: {url}")
    if keywords:
        print(f"Using keywords: {keywords}")
    print("-" * 40)
    
    try:
        # Create a temporary Python script to call the custom download function
        temp_script = f"""
import sys
sys.path.append('.')
from main import download_from_custom_site

url = "{url}"
keywords = {keywords if keywords else 'None'}

result = download_from_custom_site(url, keywords)
if result:
    print(f"SUCCESS: File downloaded to {{result}}")
else:
    print("ERROR: Download failed")
    sys.exit(1)
"""
        
        # Set environment variable to handle Unicode properly
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        
        # Run the temporary script
        result = subprocess.run([sys.executable, '-c', temp_script], 
                              capture_output=True, 
                              text=True, 
                              check=True,
                              env=env,
                              encoding='utf-8',
                              errors='replace')
        
        # Print the output (replace problematic Unicode chars)
        if result.stdout:
            stdout_clean = result.stdout.replace('✅', 'SUCCESS').replace('❌', 'ERROR').replace('📁', '').replace('📏', '')
            print(stdout_clean)
        
        if result.stderr:
            stderr_clean = result.stderr.replace('✅', 'SUCCESS').replace('❌', 'ERROR').replace('📁', '').replace('📏', '')
            print("STDERR:", stderr_clean)
        
        print("-" * 40)
        print("Custom download completed successfully!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Custom download failed with return code {e.returncode}")
        if e.stdout:
            stdout_clean = e.stdout.replace('✅', 'SUCCESS').replace('❌', 'ERROR').replace('📁', '').replace('📏', '')
            print("STDOUT:", stdout_clean)
        if e.stderr:
            stderr_clean = e.stderr.replace('✅', 'SUCCESS').replace('❌', 'ERROR').replace('📁', '').replace('📏', '')
            print("STDERR:", stderr_clean)
        return False
    except Exception as e:
        print(f"ERROR: Unexpected error during custom download: {str(e)}")
        return False

def run_hardcoded_extraction():
    """Run the hardcoded map.py extraction"""
    print("Running hardcoded extraction (map.py)...")
    print("-" * 40)
    
    try:
        # Run map.py and capture output
        result = subprocess.run([sys.executable, 'map.py'], 
                              capture_output=True, 
                              text=True, 
                              check=True)
        
        # Print the output
        if result.stdout:
            print(result.stdout)
        
        if result.stderr:
            print("STDERR:", result.stderr)
        
        print("-" * 40)
        print("Hardcoded extraction completed successfully!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Hardcoded extraction failed with return code {e.returncode}")
        if e.stdout:
            print("STDOUT:", e.stdout)
        if e.stderr:
            print("STDERR:", e.stderr)
        return False
    except Exception as e:
        print(f"ERROR: Unexpected error during hardcoded extraction: {str(e)}")
        return False

def run_universal_extraction():
    """Run the universal mapping extraction if available"""
    if not os.path.exists('universal_map.py'):
        print("Universal mapping script (universal_map.py) not found - skipping")
        return True
    
    print("Running universal extraction (universal_map.py)...")
    print("-" * 40)
    
    try:
        # Run universal_map.py and capture output
        result = subprocess.run([sys.executable, 'universal_map.py'], 
                              capture_output=True, 
                              text=True, 
                              check=True)
        
        # Print the output
        if result.stdout:
            print(result.stdout)
        
        if result.stderr:
            print("STDERR:", result.stderr)
        
        print("-" * 40)
        print("Universal extraction completed successfully!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Universal extraction failed with return code {e.returncode}")
        if e.stdout:
            print("STDOUT:", e.stdout)
        if e.stderr:
            print("STDERR:", e.stderr)
        return False
    except Exception as e:
        print(f"ERROR: Unexpected error during universal extraction: {str(e)}")
        return False

def run_comparison():
    """Run comparison scripts if available"""
    comparison_scripts = ['compare_outputs.py', 'verify_all_columns.py', 'compare_universal.py']
    
    for script in comparison_scripts:
        if os.path.exists(script):
            print(f"Running comparison script: {script}")
            print("-" * 40)
            
            try:
                result = subprocess.run([sys.executable, script], 
                                      capture_output=True, 
                                      text=True, 
                                      check=True)
                
                if result.stdout:
                    print(result.stdout)
                
                if result.stderr:
                    print("STDERR:", result.stderr)
                
                print("-" * 40)
                print(f"{script} completed successfully!")
                
            except subprocess.CalledProcessError as e:
                print(f"WARNING: {script} failed with return code {e.returncode}")
                if e.stdout:
                    print("STDOUT:", e.stdout)
                if e.stderr:
                    print("STDERR:", e.stderr)
            except Exception as e:
                print(f"WARNING: Unexpected error in {script}: {str(e)}")
            
            print()

def list_output_files():
    """List generated output files"""
    print("Generated output files:")
    print("-" * 30)

    output_files = []

    # Check both current directory and output folder
    search_locations = ['.', 'output']

    for location in search_locations:
        if os.path.exists(location):
            for file in os.listdir(location):
                if file.startswith('GFB_DATA_') and file.endswith('.xlsx'):
                    file_path = os.path.join(location, file)
                    output_files.append(file_path)

    if output_files:
        # Sort by modification time (newest first)
        output_files.sort(key=lambda f: os.path.getmtime(f), reverse=True)

        for i, file in enumerate(output_files, 1):
            mod_time = datetime.fromtimestamp(os.path.getmtime(file)).strftime('%Y-%m-%d %H:%M:%S')
            file_size = os.path.getsize(file) / 1024  # KB
            print(f"{i:2d}. {file}")
            print(f"    Created: {mod_time}, Size: {file_size:.1f} KB")
    else:
        print("No output files found.")

    print()

def main():
    """Main orchestrator function"""
    parser = argparse.ArgumentParser(description='GFB Data Processing Orchestrator')
    parser.add_argument('--download-only', action='store_true', 
                       help='Run only universal download script (main.py)')
    parser.add_argument('--hardcoded-only', action='store_true', 
                       help='Run only hardcoded extraction (map.py)')
    parser.add_argument('--universal-only', action='store_true', 
                       help='Run only universal extraction (universal_map.py)')
    parser.add_argument('--no-download', action='store_true', 
                       help='Skip download step and start with extraction')
    parser.add_argument('--no-comparison', action='store_true', 
                       help='Skip running comparison scripts')
    parser.add_argument('--list-outputs', action='store_true', 
                       help='List existing output files and exit')
    parser.add_argument('--url', type=str, 
                       help='Custom URL to download Excel file from (requires --download-only)')
    parser.add_argument('--keywords', type=str, nargs='*',
                       help='Keywords to help identify the right download (use with --url)')
    
    args = parser.parse_args()
    
    print_banner()
    
    # Handle list outputs only
    if args.list_outputs:
        list_output_files()
        return
    
    # Check dependencies and files
    if not check_dependencies():
        sys.exit(1)
    
    if not check_files():
        sys.exit(1)
    
    print("All dependencies and files found!")
    print()
    
    # Track success status
    success = True
    
    # Handle custom download with URL
    if args.url and not args.download_only:
        print("ERROR: --url requires --download-only flag")
        sys.exit(1)
    
    # Run pipeline based on arguments
    if args.download_only:
        if args.url:
            # Create temporary script to call download_from_custom_site
            success = run_custom_download(args.url, args.keywords)
        else:
            success = run_download()
    elif args.universal_only:
        success = run_universal_extraction()
    elif args.hardcoded_only:
        success = run_hardcoded_extraction()
    else:
        # Run complete pipeline (default)
        print("Running complete pipeline...")
        print()
        
        # Step 1: Download the data file (unless skipped)
        if not args.no_download:
            if args.url:
                success = run_custom_download(args.url, args.keywords)
            else:
                success = run_download()
            print()
        else:
            success = True
            print("Skipping download step...")
            print()
        
        # Step 2: Run hardcoded extraction (main solution)
        if success:
            success = run_hardcoded_extraction()
            print()
        
        # Step 3: Run universal extraction if available
        if success:
            universal_success = run_universal_extraction()
            print()
    
    # Run comparisons unless disabled
    if not args.no_comparison and success:
        print("Running comparison analysis...")
        print()
        run_comparison()
    
    # List generated files
    list_output_files()
    
    # Final status
    print("=" * 60)
    if success:
        print("PIPELINE COMPLETED SUCCESSFULLY!")
    else:
        print("PIPELINE COMPLETED WITH ERRORS!")
        sys.exit(1)
    print("=" * 60)

if __name__ == "__main__":
    main()