# HTML to PowerPoint Converter

This tool converts HTML content into PowerPoint presentations.

## Features

- Converts HTML content to PowerPoint format
- Preserves basic HTML formatting
- Supports various HTML elements and styling

## Usage

To use the HTML to PowerPoint converter, run:

```bash
./html2ppt <input_html_file> <output_pptx_file>
```

Or using the Python script:

```bash
python html2ppt.py <input_html_file> <output_pptx_file>
```

## Installation

1. Clone or download the repository:
   ```bash
   git clone https://github.com/streetquant/html2ppt.git
   cd html2ppt
   ```

2. Create a virtual environment (recommended for Python dependency management):
   ```bash
   python -m venv venv
   ```

3. Activate the virtual environment:
   ```bash
   # On Linux/Mac:
   source venv/bin/activate

   # On Windows:
   venv\Scripts\activate
   ```

4. Install the required dependencies:
   ```bash
   pip install python-pptx playwright
   ```

5. Install Playwright's browser binaries:
   ```bash
   playwright install chromium
   ```

6. Make the script executable:
   ```bash
   chmod +x html2ppt
   ```

The tool must run within the virtual environment that contains the required dependencies.

## Dependencies

- python-pptx
- playwright
- chromium browser (installed via playwright)

## Running the Tool

Remember to always activate the virtual environment before running the tool:

```bash
source venv/bin/activate  # On Linux/Mac
# or
venv\Scripts\activate     # On Windows

./html2ppt <input_html_file> <output_pptx_file>
```

The wrapper script (html2ppt) will automatically handle virtual environment activation, but if you run the Python script directly, ensure your virtual environment is active.

## License

This project is licensed under the MIT License - see the LICENSE file for details.