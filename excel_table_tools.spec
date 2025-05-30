import sys
import os
import shutil
sys.setrecursionlimit(sys.getrecursionlimit() * 5)

# -*- mode: python ; coding: utf-8 -*-

# Get the absolute path to the project root
PROJ_ROOT = os.path.abspath(os.path.dirname('excel_table_tools.py'))

# Set output directory based on platform
import platform
if platform.system() == 'Windows':
    DIST_PATH = os.path.join('GenerateExecutable', 'windows')
elif platform.system() == 'Darwin':  # macOS
    DIST_PATH = os.path.join('GenerateExecutable', 'macos')
else:  # Linux and others
    DIST_PATH = os.path.join('GenerateExecutable', 'linux')

# Function to discover all Python modules in a directory
def discover_modules(start_dir, base_package=''):
    if not os.path.exists(start_dir):
        print(f"Warning: Directory {start_dir} does not exist!")
        return []
        
    modules = []
    print(f"\nScanning directory: {start_dir}")
    
    for item in os.listdir(start_dir):
        path = os.path.join(start_dir, item)
        if os.path.isfile(path) and item.endswith('.py'):
            module_name = item[:-3]  # Remove .py extension
            if module_name != '__init__':  # Skip __init__.py
                full_module = f"{base_package}.{module_name}" if base_package else module_name
                modules.append(full_module)
                print(f"Found module: {full_module}")
        elif os.path.isdir(path) and not item.startswith(('__', '.')):
            # Recursively discover modules in subdirectories
            sub_modules = discover_modules(path, f"{base_package}.{item}" if base_package else item)
            modules.extend(sub_modules)
    return modules

# Clean up previous build artifacts
for cleanup_dir in ['build', 'dist']:
    if os.path.exists(cleanup_dir):
        print(f"Cleaning up {cleanup_dir} directory...")
        shutil.rmtree(cleanup_dir)

# Create temporary build directory with all necessary subdirectories
temp_build_dir = os.path.join('GenerateExecutable', '.build_temp')
temp_build_subdir = os.path.join(temp_build_dir, 'excel_table_tools')
for build_dir in [temp_build_dir, temp_build_subdir]:
    if not os.path.exists(build_dir):
        print(f"Creating build directory: {build_dir}")
        os.makedirs(build_dir)

# Create output directory if it doesn't exist
if not os.path.exists(DIST_PATH):
    print(f"Creating output directory: {DIST_PATH}")
    os.makedirs(DIST_PATH)

# Discover all operation modules
operations_dir = os.path.join(PROJ_ROOT, 'src', 'operations')
print(f"\nLooking for operations in: {operations_dir}")
operation_modules = discover_modules(operations_dir, 'src.operations')

print("\nDiscovered operation modules:")
for module in operation_modules:
    print(f"  - {module}")

# Base hidden imports
base_imports = [
    'pandas',
    'openpyxl',
    'tabulate',
    'src',
    'src.operations',
    'src.translations'
]

# Combine base imports with discovered operation modules
all_hidden_imports = base_imports + operation_modules

print("\nFinal hidden imports:")
for imp in all_hidden_imports:
    print(f"  - {imp}")

block_cipher = None

# Collect all data files - only include directories that exist
datas = [
    ('resources', 'resources'),
]

# Only add config directory if it exists
config_dir = os.path.join('src', 'config')
if os.path.exists(config_dir):
    datas.append((config_dir, 'src/config'))
    print(f"Including config directory: {config_dir}")
else:
    print(f"Config directory not found, skipping: {config_dir}")
    # Create an empty config directory for the build
    os.makedirs(config_dir, exist_ok=True)
    # Create a placeholder file to ensure the directory is included
    placeholder_file = os.path.join(config_dir, '.gitkeep')
    with open(placeholder_file, 'w') as f:
        f.write('')
    datas.append((config_dir, 'src/config'))
    print(f"Created empty config directory for packaging: {config_dir}")

a = Analysis(
    ['excel_table_tools.py'],
    pathex=[PROJ_ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=all_hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    workpath=temp_build_dir  # Use temporary build directory
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelTableTools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    distpath=DIST_PATH
)

# Clean up temporary build directory
print(f"\nCleaning up temporary build directory: {temp_build_dir}")
if os.path.exists(temp_build_dir):
    shutil.rmtree(temp_build_dir)

# We don't need the COLLECT step since we're using --onefile
