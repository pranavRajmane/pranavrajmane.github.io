#!/usr/bin/env python3
import sys
import subprocess
import shutil
from pathlib import Path

def find_pandoc():
    """Finds the pandoc executable in common Homebrew paths or PATH."""
    pandoc_path = shutil.which("pandoc")
    if pandoc_path:
        return pandoc_path
    
    # Fallback for common Homebrew locations if not in PATH
    for path_str in ["/opt/homebrew/bin/pandoc", "/usr/local/bin/pandoc"]:
        path_obj = Path(path_str)
        if path_obj.is_file():
            return str(path_obj)
            
    return None

def run_command(cmd_list):
    """Executes a shell command and checks for errors."""
    try:
        subprocess.run(cmd_list, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running command: {' '.join(cmd_list)}", file=sys.stderr)
        print(f"STDOUT: {e.stdout}", file=sys.stderr)
        print(f"STDERR: {e.stderr}", file=sys.stderr)
        return False
    return True

def convert_docx_to_md(docx_file_path):
    """
    Converts a .docx file to a .md file in the same directory
    and extracts all images into a local assets folder.
    """
    
    # 1. Validate input and find Pandoc
    docx_file = Path(docx_file_path)
    if not docx_file.exists() or docx_file.suffix != '.docx':
        print(f"Error: File not found or is not a .docx file: {docx_file}", file=sys.stderr)
        return

    pandoc_exec = find_pandoc()
    if not pandoc_exec:
        print("Error: Pandoc executable not found.", file=sys.stderr)
        print("Please install pandoc with 'brew install pandoc'", file=sys.stderr)
        return

    # 2. Define file paths
    final_md_file = docx_file.with_suffix('.md')
    
    # Define the folder where images will be extracted
    # This creates a Jekyll-friendly path like: .../assets/images/My-Blog-Post/
    post_assets_dir = docx_file.parent / 'assets' / 'images' / docx_file.stem
    # Create this directory if it doesn't exist
    post_assets_dir.mkdir(parents=True, exist_ok=True)

    print(f"Converting '{docx_file.name}' to Markdown...")

    # 3. Run Pandoc to convert DOCX -> MD and extract images
    pandoc_command = [
        pandoc_exec,
        '-f', 'docx',      # from format: docx
        '-t', 'markdown',  # to format: markdown
        '--extract-media', str(post_assets_dir), # Extract images to this folder
        str(docx_file),    # input file
        '-o', str(final_md_file) # output file
    ]
    
    if not run_command(pandoc_command):
        print("Failed to convert .docx to .md with Pandoc.", file=sys.stderr)
        return

    print(f"\n✅ Success! Created:")
    print(f"  📄 Markdown: {final_md_file.name}")
    # Note: Pandoc creates a 'media' subfolder inside your target dir
    print(f"  🖼️ Images:  {post_assets_dir.relative_to(docx_file.parent)}/media")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 /path/to/docx2md.py /path/to/your/document.docx", file=sys.stderr)
        sys.exit(1)
        
    convert_docx_to_md(sys.argv[1])
