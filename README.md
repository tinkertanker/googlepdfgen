# Generate PDFs from Google Sheets and Google Slides/PowerPoint

This script takes a Google Slides or PowerPoint template and generates PDFs from
them by switching out placeholders based on a Google Sheets document. This is
ideal for applications such as mass certificate generation, where you can
later conveniently send out the PDFs using the mail-merge software of your
choice.

It also linearises all PDFs generated for web viewing, downscales images, and
optimises the document for printing. Additionally, all the PDFs are PDF/A-2b
compliant.

## Usage

The Google Sheet needs to be in the following format:

| filename | \<name> | \<class> | file                                   |
| :------- | :------ | :------: | :------------------------------------- |
| a_file   | Ivan    |   100    | [this will be filled in by the script] |

This generates `a_file.pdf` where `<name>` would be replaced by `Ivan` and
`<class>` by `100`. The "file" column would be filled in with the Google Drive
link. The script will ignore columns that do not match `^(filename|file|<.+>)$`.

You also need a `client_secret.json` file in the root directory for this to
work. To get one, refer to <https://stackoverflow.com/a/55416898>.

```bash
# Install LibreOffice for rendering the template slides as PDFs (yes this is necessary)
brew install libreoffice # macOS (get homebrew from https://brew.sh/)
sudo apt install libreoffice # Linux (apt)
sudo yum install libreoffice # Linux (yum)

# Install Ghostscript for cleaning up the PDFs (yes this is necessary)
brew install ghostscript # macOS (get homebrew from https://brew.sh/)
sudo apt install ghostscript # Linux (apt)
sudo yum install ghostscript # Linux (yum)

# Install Python dependencies
python3 -m pip install -U pipenv  # if you don't already have it installed
pipenv install

# Run script
pipenv run python gen.py --sheet X --template X --output X

# For more information about the arguments you can pass, check out the help text (please do this before asking for help)
pipenv run python gen.py --help
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
