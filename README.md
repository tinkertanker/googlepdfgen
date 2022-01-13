# Generate PDFs from Google Slides & Sheets

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

# For more information about the arguments you can pass, check out the help text
pipenv run python gen.py --help
```
