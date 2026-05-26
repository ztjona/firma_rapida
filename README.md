# Firma Rápida

Quick PDF digital signing tool for Windows. Sign PDFs (and Word documents) with your `.p12`/`.pfx` certificate through a visual interface — no command line required.

## Download

Go to [Releases](../../releases) and download the latest `FirmaRapida.exe`. Double-click to run — **no installation needed**.

> **Note:** To sign Word (`.docx`) files, Microsoft Word or LibreOffice must be installed on your machine.

## Features

- **Visual signature placement** — preview the PDF, drag signature boxes to the exact position you want
- **Multiple signatures** — place signatures on different pages or positions in the same document
- **Remembers your certificate** — certificate path and password are stored encrypted locally; no need to re-enter them every time
- **Word support** — open `.docx` files directly; they are auto-converted to PDF before signing
- **Page navigation and zoom** — browse multi-page documents before signing
- **Configurable metadata** — set signer name, location, reason, and signature box size

## Usage

1. Run `FirmaRapida.exe`
2. On first run, configure your certificate (`.p12` / `.pfx`) and password via **Configuración**
3. Open a PDF or Word document
4. Click **Agregar firma** and drag the signature box to the desired position
5. Repeat for additional signatures on other pages if needed
6. Click **Firmar** — the signed PDF is saved next to the original

## Running from source

**Requirements:** Python 3.9+

```bash
pip install -r requirements.txt
python firma_rapida.py
```

Or run `instalar.bat` to install dependencies automatically, then:

```bash
python firma_rapida.py
```

## Building the executable locally

```bat
build.bat
```

The output will be at `dist\FirmaRapida.exe`. Requires Python and pip installed on your machine.

## Releasing a new version

```bash
git tag v1.3
git push origin main --tags
```

GitHub Actions will build and publish `FirmaRapida.exe` to the Releases page automatically.

## Security note

Your certificate password is stored **locally only**, encrypted with a machine-specific key using the `cryptography` library (Fernet). It is never transmitted anywhere.
