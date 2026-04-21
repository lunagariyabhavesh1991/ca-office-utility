import os
import subprocess
import shutil
import sys
import json
import zipfile
from datetime import datetime

def run_command(command, description):
    print(f"\n--- {description} ---")
    print(f"Running: {command}")
    result = subprocess.run(command, shell=True)
    if result.returncode != 0:
        print(f"Error: {description} failed.")
        sys.exit(1)
    return True

def generate_build_metadata(project_dir, version="v1.1.2"):
    """Generates assets/build_info.json with current build timestamp"""
    assets_dir = os.path.join(project_dir, "assets")
    os.makedirs(assets_dir, exist_ok=True)
    
    # We use local time for the build stamp to help user identify their specific build
    build_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    metadata = {
        "version": version,
        "build_time": build_time,
        "environment": "Production"
    }
    
    metadata_path = os.path.join(assets_dir, "build_info.json")
    with open(metadata_path, 'w') as f:
        json.dump(metadata, f, indent=4)
    
    print(f"Build Metadata Generated: {version} ({build_time})")
    return metadata

def create_secure_zip(source_dir, output_zip):
    """Creates a clean ZIP for distribution, excluding source files and internal tools"""
    print(f"Creating secure ZIP: {output_zip}")
    
    # Files/folders to exclude from the ZIP
    exclude_extensions = {'.py', '.spec', '.bat', '.pyc', '.git', '.github'}
    exclude_folders = {'__pycache__', 'obfuscated_src', 'build', 'scripts', 'tests', 'deployments'}

    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(source_dir):
            # Filter folders
            dirs[:] = [d for d in dirs if d not in exclude_folders]
            
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in exclude_extensions:
                    continue
                
                # Full path of file
                file_path = os.path.join(root, file)
                
                # Relative path for the ZIP (should be relative to source_dir's parent to include the folder name)
                # Or relative to source_dir if we want files flat inside the zip root
                arcname = os.path.relpath(file_path, source_dir)
                zipf.write(file_path, arcname)
    
    print(f"Successfully created: {output_zip} ({os.path.getsize(output_zip) / (1024*1024):.2f} MB)")

def main():
    # Check for --standard flag
    is_standard = "--standard" in sys.argv
    version = "v1.1.2"
    
    project_dir = os.getcwd()
    dist_dir = os.path.join(project_dir, "dist")
    obf_dir = os.path.join(project_dir, "obfuscated_src")
    
    # 0. Generate Metadata
    generate_build_metadata(project_dir, version)

    if is_standard:
        print("!!! PREPARING STANDARD VERSION BUILD (EXCLUDING PREMIUM MODULES) !!!")
        # Temporarily modify app_window.py to set IS_PREMIUM_BUILD = False
        app_win_path = os.path.join(project_dir, "ui", "app_window.py")
        with open(app_win_path, "r", encoding="utf-8") as f:
            content = f.read()
        with open(app_win_path, "w", encoding="utf-8") as f:
            f.write(content.replace("IS_PREMIUM_BUILD = True", "IS_PREMIUM_BUILD = False"))

    # 1. Clean previous builds
    print("Cleaning previous builds...")
    for d in ["build", "dist", "obfuscated_src"]:
        if os.path.exists(d):
            try:
                shutil.rmtree(d)
                print(f"Cleaned {d}")
            except Exception as e:
                print(f"Warning: Could not clean {d} - {e}")
                if os.name == 'nt':
                    os.system(f'rd /s /q "{d}"')

    # 2. Obfuscate source code (Preparation)
    print("Preparing source for build (Staging)...")
    py_cmd = sys.executable
    os.makedirs(obf_dir, exist_ok=True)

    # Base targets
    targets = ["main.py", "core", "utils", "ui", "assets", "services"]
    
    for target in targets:
        if not os.path.exists(target): continue
        if os.path.isdir(target):
            # Special handling for standard build exclusions
            if is_standard and target == "ui":
                shutil.copytree(target, os.path.join(obf_dir, target), dirs_exist_ok=True,
                               ignore=shutil.ignore_patterns('gst_pack_view.py', 'cma_dpr_builder_view.py', 'invoice_parser_view.py'))
            elif is_standard and target == "services":
                # Only copy standard services
                os.makedirs(os.path.join(obf_dir, "services"), exist_ok=True)
                standard_services = ["__init__.py", "bank_parsers.py", "bank_parser_base.py", 
                                   "bank_statement_parser.py", "excel_csv_exporter.py", 
                                   "pdf_table_extractor.py", "tally_xml_exporter.py", "tally_xml_generator.py"]
                for s in standard_services:
                    s_path = os.path.join("services", s)
                    if os.path.exists(s_path):
                        shutil.copy(s_path, os.path.join(obf_dir, s_path))
            else:
                shutil.copytree(target, os.path.join(obf_dir, target), dirs_exist_ok=True)
        else:
            shutil.copy(target, os.path.join(obf_dir, target))

    # Copy spec file and other essentials
    for f in ["CA_Office_PDF_Utility.spec"]:
        if os.path.exists(f):
            shutil.copy(f, os.path.join(obf_dir, f))

    # 4. Run PyInstaller from the obfuscated source directory
    print("\nRunning PyInstaller...")
    os.chdir(obf_dir)
    
    if is_standard:
        # Patch the SPEC file to include exclusions
        spec_path = "CA_Office_PDF_Utility.spec"
        exclusions = [
            "ui.views.gst_pack_view", "ui.views.cma_dpr_builder_view", "ui.views.invoice_parser_view",
            "services.gst_pack", "services.cma", "services.invoice_parser"
        ]
        with open(spec_path, "r", encoding="utf-8") as f:
            spec_content = f.read()
        
        import re
        exclude_list_str = ", ".join([f"'{m}'" for m in exclusions])
        spec_content = re.sub(r"excludes=\[([^\]]*)\]", f"excludes=[\\1, {exclude_list_str}]", spec_content)
        
        with open(spec_path, "w", encoding="utf-8") as f:
            f.write(spec_content)

    pyinstaller_cmd = f'"{py_cmd}" -m PyInstaller -y CA_Office_PDF_Utility.spec'
    run_command(pyinstaller_cmd, "PyInstaller Build")
    
    # 5. Move final build back to root dist
    os.chdir(project_dir)
    
    # Revert app_window.py if it was changed
    if is_standard:
        app_win_path = os.path.join(project_dir, "ui", "app_window.py")
        with open(app_win_path, "r", encoding="utf-8") as f:
            content = f.read()
        with open(app_win_path, "w", encoding="utf-8") as f:
            f.write(content.replace("IS_PREMIUM_BUILD = False", "IS_PREMIUM_BUILD = True"))

    output_subfolder = "BKL_Office_Standard" if is_standard else "CA_Office_PDF_Utility"
    final_output = os.path.join(obf_dir, "dist", "CA_Office_PDF_Utility")
    dest_path = os.path.join(dist_dir, output_subfolder)
    
    if os.path.exists(final_output):
        print(f"Moving results to: {dest_path}")
        if os.path.exists(dest_path): shutil.rmtree(dest_path)
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        shutil.move(final_output, dest_path)
        
        # 6. Create Secure Sanitized ZIP
        platform_suffix = "Windows" if os.name == 'nt' else "macOS"
        zip_name = f"CA_Office_Utility_{platform_suffix}_{version}.zip"
        if is_standard: zip_name = f"BKL_Office_Standard_{platform_suffix}_{version}.zip"
        
        zip_path = os.path.join(dist_dir, zip_name)
        create_secure_zip(dest_path, zip_path)

        # 6b. Automation: Dummy file for legacy workflow compatibility
        legacy_name = f"CA_Office_PDF_Utility_{version}.zip"
        shutil.copy(zip_path, os.path.join(dist_dir, legacy_name))
        
        # 7. CI-ONLY: Direct Upload to GitHub Release (Bypassing CLI limits)
        if os.getenv("GITHUB_ACTIONS") == "true":
            gh_token = os.getenv("PUBLIC_REPO_TOKEN")
            if gh_token:
                print(f"CI Mode: Attempting direct upload of {zip_name}...")
                try:
                    import requests
                    repo = "marutitechsolutions/ca-office-utility"
                    tag = version
                    # 1. Get Release ID
                    r = requests.get(f"https://api.github.com/repos/{repo}/releases/tags/{tag}", 
                                     headers={"Authorization": f"token {gh_token}"})
                    if r.status_code == 200:
                        release_id = r.json()["id"]
                        # 2. Upload Asset
                        with open(zip_path, "rb") as f:
                            upload_url = f"https://uploads.github.com/repos/{repo}/releases/{release_id}/assets?name={zip_name}"
                            r_up = requests.post(upload_url, headers={
                                "Authorization": f"token {gh_token}",
                                "Content-Type": "application/zip"
                            }, data=f)
                            if r_up.status_code in [201, 200]:
                                print(f"Successfully uploaded {zip_name} via API!")
                            else:
                                print(f"API Upload failed: {r_up.status_code} - {r_up.text}")
                except Exception as e:
                    print(f"CI Upload Error: {e}")

        print(f"\nSUCCESS! Build complete.")
        print(f"Build Folder: {dest_path}")
        print(f"Release ZIP: {zip_path}")
    else:
        print(f"Error: Build folder not found at {final_output}")

if __name__ == "__main__":
    main()
