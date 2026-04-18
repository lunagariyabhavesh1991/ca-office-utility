import os
import subprocess
import shutil
import sys

def run_command(command, description):
    print(f"--- {description} ---")
    print(f"Running: {command}")
    result = subprocess.run(command, shell=True)
    if result.returncode != 0:
        print(f"Error: {description} failed.")
        sys.exit(1)

def main():
    # Check for --standard flag
    is_standard = "--standard" in sys.argv
    
    project_dir = os.getcwd()
    dist_dir = os.path.join(project_dir, "dist")
    obf_dir = os.path.join(project_dir, "obfuscated_src")
    
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
                # Fallback for Windows if file is locked
                if os.name == 'nt':
                    os.system(f'rd /s /q "{d}"')

    # 2. Obfuscate source code (Stay within Trial limit)
    print("Obfuscating core logic source code with PyArmor...")
    # Use system default python
    py_cmd = sys.executable
    
    # Base files to obfuscate - Check if we are in a 'src' structure
    base_prefix = "" if os.path.exists("main.py") else ("src/" if os.path.exists("src/main.py") else "")
    targets = [f"{base_prefix}main.py", f"{base_prefix}core", f"{base_prefix}utils"]
    
    # Files to move after obfuscation if they get flattened
    move_after = []

    if is_standard:
        # We manually list the basic services to avoid the premium ones
        service_files = [
            "__init__.py", "bank_parsers.py", "bank_parser_base.py", 
            "bank_statement_parser.py", "excel_csv_exporter.py", 
            "pdf_table_extractor.py", "tally_xml_exporter.py", "tally_xml_generator.py"
        ]
        for f in service_files:
            targets.append(os.path.join(base_prefix, "services", f))
            move_after.append(f)
    else:
        targets.append("services")

    targets_str = " ".join(targets)
    # run_command(f'{py_cmd} -m pyarmor.cli gen -O "{obf_dir}" {targets_str}', 
    #             "Core Obfuscation")
    
    # Fallback: Copy source directly if obfuscation is skipped/fails
    print("Skipping obfuscation due to license limits. Copying raw source for build...")
    os.makedirs(obf_dir, exist_ok=True)
    for target in targets:
        if os.path.isdir(target):
            shutil.copytree(target, os.path.join(obf_dir, target), dirs_exist_ok=True)
        else:
            dest_path = os.path.join(obf_dir, target)
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            shutil.copy(target, dest_path)

    # FIX: If PyArmor flattened the services files into obfuscated_src, move them to services/
    if is_standard:
        dest_services = os.path.join(obf_dir, "services")
        os.makedirs(dest_services, exist_ok=True)
        for f in move_after:
            src = os.path.join(obf_dir, f)
            if os.path.exists(src):
                dest = os.path.join(dest_services, f)
                if os.path.exists(dest): os.remove(dest)
                shutil.move(src, dest)

    # 3. Copy other files and assets to obfuscated source directory
    print("Copying remaining files to build directory...")
    # List of folders to copy normally (non-obfuscated)
    folders_to_copy = ["ui", "assets"]
    
    for item in folders_to_copy:
        if os.path.exists(item):
            # If standard, don't copy the premium view files in ui/views
            if is_standard and item == "ui":
                shutil.copytree(item, os.path.join(obf_dir, item), dirs_exist_ok=True,
                               ignore=shutil.ignore_patterns('gst_pack_view.py', 'cma_dpr_builder_view.py', 'invoice_parser_view.py'))
            else:
                shutil.copytree(item, os.path.join(obf_dir, item), dirs_exist_ok=True)
    
    # Copy spec file and icon
    for f in ["CA_Office_PDF_Utility.spec", "assets/logo.png"]:
        if os.path.exists(f):
            dest_f = os.path.join(obf_dir, f)
            os.makedirs(os.path.dirname(dest_f), exist_ok=True)
            shutil.copy(f, dest_f)

    # Safety Check: Explicitly remove sensitive folders from obf_dir to prevent packaging
    sensitive_folders = ["scripts", "security", "internal_docs", "tests", "deployments"]
    for sf in sensitive_folders:
        path = os.path.join(obf_dir, sf)
        if os.path.exists(path):
            print(f"!!! SECURITY REMOVAL: Deleting {sf} from build staging !!!")
            shutil.rmtree(path)

    # 4. Run PyInstaller from the obfuscated source directory
    print("Running PyInstaller on protected code...")
    os.chdir(obf_dir)
    
    if is_standard:
        # Patch the SPEC file to include exclusions
        spec_path = "CA_Office_PDF_Utility.spec"
        exclusions = [
            "ui.views.gst_pack_view",
            "ui.views.cma_dpr_builder_view",
            "ui.views.invoice_parser_view",
            "services.gst_pack",
            "services.cma",
            "services.invoice_parser"
        ]
        with open(spec_path, "r", encoding="utf-8") as f:
            spec_content = f.read()
        
        # Inject into the excludes list
        exclude_list_str = ", ".join([f"'{m}'" for m in exclusions])
        import re
        spec_content = re.sub(r"excludes=\[([^\]]*)\]", f"excludes=[\\1, {exclude_list_str}]", spec_content)
        
        with open(spec_path, "w", encoding="utf-8") as f:
            f.write(spec_content)

    pyinstaller_cmd = f"{py_cmd} -m PyInstaller -y CA_Office_PDF_Utility.spec"
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
        
        try:
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            shutil.move(final_output, dest_path)
        except Exception as e:
            print(f"Standard move failed: {e}. Trying system move.")
            if os.name == 'nt':
                os.system(f'robocopy "{final_output}" "{dest_path}" /E /MOVE /NP /R:3 /W:1')
            else:
                os.system(f'cp -R "{final_output}" "{dest_path}" && rm -rf "{final_output}"')
        
        if os.path.exists(os.path.join(dest_path, "CA_Office_PDF_Utility.exe")) or os.path.exists(os.path.join(dest_path, "CA_Office_PDF_Utility")):
            print(f"\nSUCCESS! Build complete in dist/{output_subfolder}")
        else:
            print("Warning: Move completed but executable not found in destination.")
    else:
        print(f"Error: Build folder not found at {final_output}")

if __name__ == "__main__":
    main()
