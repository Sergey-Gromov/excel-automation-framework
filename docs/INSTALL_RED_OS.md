# Installing the 513 Excel Transformation on RED OS

This guide describes how to install the `513-finish` Excel transformation project on another RED OS Linux computer.

The production batch command is:

```bash
python run_batch.py incoming transformed
```

The default validation sample is `samples/input_structural.xlsx` (the production
sample truncated to 100 rows), so the plain command is correct and fast for the
6969-row production workbooks.

Important rule: **do not copy the `.venv` folder from the old computer**. Create a fresh virtual environment on the RED OS machine.

---

## Required folder layout

If you copy the local `excel-framework-core` folder together with this project, place it **next to** `513-finish`, not inside it.

Required structure:

```text
~/excel-513/
├── 513-finish/
└── excel-framework-core/
```

Example:

```text
/home/USER/excel-513/513-finish
/home/USER/excel-513/excel-framework-core
```

The folder name must be exactly:

```text
excel-framework-core
```

This layout is required for the command:

```bash
pip install -e ../excel-framework-core
```

That command must be run from inside:

```bash
~/excel-513/513-finish
```

---

# 1. Create a portable archive on the current computer

From the parent folder containing `513-finish`:

```bash
cd "/home/sergio/MOEK-local/Ликвидность/foo513_dev"

tar \
  --exclude='.venv' \
  --exclude='.git' \
  --exclude='reports' \
  --exclude='needs_review' \
  --exclude='transformed' \
  --exclude='incoming' \
  --exclude='__pycache__' \
  --exclude='*.pyc' \
  --exclude='.~lock*' \
  -czf 513-finish-redos.tar.gz \
  513-finish
```

This creates:

```bash
513-finish-redos.tar.gz
```

If you are using a local `excel-framework-core` folder, transfer that folder separately and place it next to `513-finish` on the RED OS computer.

---

# 2. Transfer files to RED OS

Example using `scp`:

```bash
scp 513-finish-redos.tar.gz USER@REDOS-PC:/home/USER/
```

Replace:

```text
USER
REDOS-PC
```

with the actual username and computer name or IP address.

If using a local `excel-framework-core` folder, copy it as well:

```bash
scp -r excel-framework-core USER@REDOS-PC:/home/USER/
```

You can also copy the files by USB.

---

# 3. Install system packages on RED OS

For modern RED OS with `dnf`:

```bash
sudo dnf update -y
sudo dnf install -y git python3 python3-pip python3-venv
```

For older RED OS using `yum`:

```bash
sudo yum update -y
sudo yum install -y git python3 python3-pip
```

---

# 4. Check the Python version

```bash
python3 --version
```

The project requires:

```text
Python 3.10 or newer
```

Recommended:

```text
Python 3.11
```

If the system Python is older than 3.10, try installing Python 3.11:

```bash
sudo dnf install -y python3.11 python3.11-pip python3.11-venv
```

If that package is not available, try:

```bash
sudo dnf module install -y python311
```

Then check:

```bash
python3.11 --version
```

---

# 5. Unpack the project

Create the target parent folder:

```bash
mkdir -p ~/excel-513
```

Unpack the project archive:

```bash
tar -xzf ~/513-finish-redos.tar.gz -C ~/excel-513
```

If you copied `excel-framework-core` to `/home/USER/`, move it into the same parent folder:

```bash
mv ~/excel-framework-core ~/excel-513/excel-framework-core
```

The final structure should be:

```text
~/excel-513/
├── 513-finish/
└── excel-framework-core/
```

Enter the project folder:

```bash
cd ~/excel-513/513-finish
```

Verify that required files exist:

```bash
ls transform.py run_batch.py run_production.py sh513.xlsx samples/input_structural.xlsx
```

If `sh513.xlsx` or `samples/input_structural.xlsx` are missing, copy them
separately from the original computer. (Regenerate the structural sample from
`samples/input_big.xlsx` with `python samples/make_structural_sample.py`.)

---

# 6. Create a fresh virtual environment

If your default `python3` is already Python 3.10 or newer:

```bash
python3 -m venv .venv
```

If you installed Python 3.11 separately:

```bash
/usr/bin/python3.11 -m venv .venv
```

Activate the virtual environment:

```bash
source .venv/bin/activate
```

Upgrade pip:

```bash
python -m pip install --upgrade pip
```

Verify that the active Python is the project virtual environment Python:

```bash
which python
```

Expected example:

```bash
/home/USER/excel-513/513-finish/.venv/bin/python
```

---

# 7. Install Python dependencies

Choose one of the options below.

---

## Option A: Install from a local `excel-framework-core` folder

Use this option if you copied the `excel-framework-core` folder together with the project.

From inside:

```bash
cd ~/excel-513/513-finish
```

run:

```bash
pip install "openpyxl>=3.1.0" "pytest>=8.0"
pip install -e ../excel-framework-core
```

This is the option that matches the required folder layout:

```text
~/excel-513/
├── 513-finish/
└── excel-framework-core/
```

Do not use `pip install -r requirements.txt` for this option if the RED OS computer cannot access the private GitHub repository, because `requirements.txt` points to GitHub.

---

## Option B: Install `excel-framework-core` from GitHub

Use this option if the RED OS computer can access the GitHub repository.

First test access:

```bash
git ls-remote https://github.com/Sergey-Gromov/excel-framework-core.git
```

If access works, install dependencies with:

```bash
pip install -r requirements.txt
```

If the repository is private, use a GitHub token:

```bash
export GITHUB_TOKEN='PASTE_TOKEN_HERE'

pip install "git+https://x-access-token:${GITHUB_TOKEN}@github.com/Sergey-Gromov/excel-framework-core.git@v0.1.0"
pip install "openpyxl>=3.1.0" "pytest>=8.0"

unset GITHUB_TOKEN
```

Do not store the token permanently in `requirements.txt`.

---

# 8. Verify the installation

Still inside the activated virtual environment:

```bash
python -c "import excel_framework, openpyxl; print('dependencies ok')"
```

Expected output:

```text
dependencies ok
```

Run the test suite:

```bash
pytest -q
```

Expected result:

```text
32 passed
```

---

# 9. Prepare production folders

Inside the project folder:

```bash
cd ~/excel-513/513-finish

mkdir -p incoming transformed needs_review reports
```

Place incoming workbooks in:

```bash
incoming/
```

Example:

```bash
cp "/path/to/5.13 ВАО ликвидность на 01.07.2026.xlsx" incoming/
```

The date source file must be in the project root:

```bash
sh513.xlsx
```

The default validation sample must be here:

```bash
samples/input_structural.xlsx
```

(Optionally keep `samples/input_big.xlsx` if you want to validate against the
full production sample with `--sample samples/input_big.xlsx`.)

---

# 10. Run the batch transformation

Activate the virtual environment if needed:

```bash
source .venv/bin/activate
```

Run:

```bash
python run_batch.py incoming transformed
```

Successful output should look like:

```text
Batch complete: 1 transformed, 0 rejected
```

Transformed files are written to:

```bash
transformed/
```

Reports are written to:

```bash
reports/
```

Rejected files, if any, are moved to:

```bash
needs_review/
```

---

# 11. Optional: run one file manually

Example:

```bash
python run_production.py \
  "incoming/5.13 ВАО ликвидность на 01.07.2026.xlsx" \
  "transformed/5.13 ВАО ликвидность на 01.07.2026.xlsx"
```

Exit code:

```text
0
```

means success.

Exit code:

```text
1
```

means the file was rejected and moved to:

```bash
needs_review/
```

---

# 12. Optional: create a convenient shell command

Add this function to `~/.bashrc`:

```bash
cat >> ~/.bashrc <<'EOF'

excel513() {
  cd "$HOME/excel-513/513-finish" || return
  source .venv/bin/activate
}
EOF
```

Apply the change:

```bash
source ~/.bashrc
```

Now you can run:

```bash
excel513
python run_batch.py incoming transformed
```

---

# Troubleshooting

## `python: command not found` after activation

Run:

```bash
hash -r
which python
```

If the result is not inside:

```bash
.venv/bin/python
```

reactivate the environment:

```bash
deactivate
source .venv/bin/activate
```

---

## `ModuleNotFoundError: No module named 'excel_framework'`

Check that the active Python is the project virtual environment Python:

```bash
which python
```

It must point to:

```bash
~/excel-513/513-finish/.venv/bin/python
```

Then reinstall the core package.

For local installation:

```bash
pip install -e ../excel-framework-core
```

For GitHub installation:

```bash
pip install -r requirements.txt
```

---

## `pip install -r requirements.txt` fails with GitHub 404

The `excel-framework-core` repository is private.

Use a GitHub token:

```bash
export GITHUB_TOKEN='PASTE_TOKEN_HERE'

pip install "git+https://x-access-token:${GITHUB_TOKEN}@github.com/Sergey-Gromov/excel-framework-core.git@v0.1.0"
pip install "openpyxl>=3.1.0" "pytest>=8.0"

unset GITHUB_TOKEN
```

Or use the local folder option:

```bash
pip install -e ../excel-framework-core
```

---

## Python version is too old

If:

```bash
python3 --version
```

shows Python 3.6, 3.8, or 3.9, install Python 3.11:

```bash
sudo dnf install -y python3.11 python3.11-pip python3.11-venv
```

or:

```bash
sudo dnf module install -y python311
```

Then recreate the virtual environment:

```bash
rm -rf .venv
/usr/bin/python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
```

Then install dependencies again.

---

## `dnf: command not found`

Older RED OS may use `yum`:

```bash
sudo yum install -y git python3 python3-pip
```

---

## Cyrillic file names or spaces cause problems

Always quote paths containing spaces or Cyrillic text:

```bash
python run_production.py \
  "incoming/5.13 ВАО ликвидность на 01.07.2026.xlsx" \
  "transformed/5.13 ВАО ликвидность на 01.07.2026.xlsx"
```
