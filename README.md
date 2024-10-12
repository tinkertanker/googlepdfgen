# Generate PDFs from Google Sheets and Google Slides/PowerPoint

This script takes a Google Slides or PowerPoint template and generates PDFs from
them by switching out placeholders based on a Google Sheets document. This is
ideal for applications such as mass certificate generation, where you can
later conveniently send out the PDFs using the mail-merge software of your
choice.

It also linearises all PDFs generated for web viewing, downscales images, and
optimises the document for printing. Additionally, all the PDFs are PDF/A-2b
compliant.

## Google Cloud Setup

Before using this script, you need to set up a Google Cloud project, enable the necessary APIs, and get a mysterious-sounding file called `client_secrets.json`. (TT staff, refer to the Notion, under "Cert Generator for Google Workspace", for a downloadable file.)

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the following APIs for your project:
   - Google Drive API
   - Google Sheets API
   - Google Slides API

To enable each API:
1. Go to "APIs & Services" > "Library"
2. Search for the API
3. Click on the API name
4. Click the "Enable" button

After enabling the APIs, you need to create credentials:

1. Go to "APIs & Services" > "Credentials"
2. Click "Create Credentials" > "OAuth client ID"
3. Select "Desktop app" as the application type
4. Download the client configuration and save it as `client_secrets.json` in the root directory of this project

Make sure the redirect URIs in your OAuth 2.0 Client ID include:
- `http://localhost:8080`
- `http://localhost:8080/` (with trailing slash)

## Usage

The Google Sheet should be structured as follows:

| filename | \<name> | \<class> | file                                   |
| :------- | :------ | :------: | :------------------------------------- |
| a_file   | Ivan    |   100    | [this will be filled in by the script] |

This sheet format uses a token system for dynamic content replacement:

1. **Tokens**: Columns with headers enclosed in angle brackets (e.g., `<name>`, `<class>`) are considered tokens.
2. **Template Placeholders**: Your Google Slides or PowerPoint template should contain placeholders that match these tokens exactly (e.g., `<name>`, `<class>`).
3. **Replacement Process**: The script processes each row, replacing the tokens in the template with the corresponding values from the sheet.

For example, using the row above:
- It generates `a_file.pdf`
- `<name>` in the template is replaced with "Ivan"
- `<class>` is replaced with "100"

The "file" column is reserved and will be populated by the script with the Google Drive link to the generated PDF.

**Note**: The script only processes columns that match the pattern `^(filename|file|<.+>)$`. Any other columns will be ignored.

## Configuration

You can now use a YAML configuration file to specify the main parameters needed for the script. This allows for easier reuse and sharing of configurations.

### Using a YAML Configuration File

1. Create a `config.yaml` file in the root directory of the project.
2. Specify the following parameters in the YAML file:

```yaml
sheet: "https://docs.google.com/spreadsheets/d/xxxxx/edit#gid=0"
template: "https://docs.google.com/presentation/d/xxxxx/edit"
output: "https://drive.google.com/drive/u/0/folders/xxxxx"
```

3. Run the script without specifying command-line arguments:

```bash
python gen.py
```

The script will automatically use the values from the `config.yaml` file.

Note: An example configuration file named `config.example.yaml` is provided in the repository. You can copy and adapt this file for your own use. The `config.yaml` file will not be uploaded to version control.

### Command-line Arguments

You can still use command-line arguments, which will override the values in the `config.yaml` file:

```bash
python gen.py --sheet X --template Y --output Z --ppi 300 --libreoffice /path/to/libreoffice --gs /path/to/ghostscript
```

If you provide any of the main arguments (sheet, template, or output) via the command line, the script will use those instead of the values in the `config.yaml` file.

For more information about the arguments you can pass, check out the help text:

```bash
python gen.py --help
```

# Install dependencies

## Install LibreOffice
LibreOffice is necessary for rendering the template slides as PDFs.

For macOS:
```bash
brew install libreoffice
```

For Linux (apt-based distributions):
```bash
sudo apt install libreoffice
```

For Linux (yum-based distributions):
```bash
sudo yum install libreoffice
```

## Install Ghostscript
Ghostscript is necessary for cleaning up the PDFs.

For macOS:
```bash
brew install ghostscript
```

For Linux (apt-based distributions):
```bash
sudo apt install ghostscript
```

For Linux (yum-based distributions):
```bash
sudo yum install ghostscript
```

## Install Python dependencies
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
pip install -r requirements.txt
```

# Run script
```bash
python gen.py --sheet X --template X --output X

# For more information about the arguments you can pass, check out the help text
python gen.py --help
```

## Why do I need to install LibreOffice for this to work?

Due to the nature of PDFs being designed for rendering and not editing, it would
be difficult to get text-replacement in PDFs to work consistently. I also could
not get SVGs to work well with image compression and font embedding as it was
pretty sketchy. Since most people prefer to generate such documents from
presentations, the most optimal local file format for doing this would be
PowerPoint. From there, I needed to render a PowerPoint file as a PDF, which is
surprisingly hard to do. The best option I had that ensures cross-platform
compatibility would be to use the LibreOffice rendering engine, which would be
most easy to use via the command line tool, and that only comes with the entire
LibreOffice installation.

## Credits

- Ivan Tung, intern, Jan 2022
- Cursor AI
