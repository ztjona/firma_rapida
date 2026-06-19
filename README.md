# Firma Rápida

Quick PDF digital signing tool for Windows. Sign PDFs (and Word documents) with your `.p12`/`.pfx` certificate through a visual interface — no command line required.

## Download

Go to [Releases](../../releases) and download the latest `FirmaRapida.exe`. Double-click to run — **no installation needed**.

> **Note:** To sign Word (`.docx`) files, Microsoft Word or LibreOffice must be installed on your machine.

## Features

- **Visual signature placement** — preview the PDF, drag signature boxes to the exact position you want
- **Multiple signatures** — place signatures on different pages or positions in the same document
- **Replicate to all pages** — place a signature once and apply it to every page in one click
- **Batch signing** — load multiple PDFs/Word files into a queue and sign them all in one shot; each signed file is saved next to its original
- **Remembers your certificate** — certificate path and password are stored encrypted locally; no need to re-enter them every time
- **Word support** — open `.docx` files directly; they are auto-converted to PDF before signing
- **Page navigation and zoom** — browse multi-page documents before signing
- **Configurable metadata** — set signer name, location, reason, and signature box size

## Usage

### Sign a single document

1. Run `FirmaRapida.exe`
2. On first run, configure your certificate (`.p12` / `.pfx`) and password via **Configuración**
3. Click **Abrir PDF/Word** and choose your document
4. Click **+ Agregar Firma** and click on the document to place the signature box; drag to reposition
5. Repeat for additional pages if needed
6. Click **FIRMAR** — choose where to save the signed PDF

### Sign all pages of a document

1. Open the document and place a signature on one page
2. Click **Replicar en todas las páginas** — the signature is copied to every page
3. Click **FIRMAR**

### Sign multiple documents at once (batch)

1. Click **Abrir PDF/Word** to open the first (reference) document
2. Click **+ Agregar archivos** in the *Documentos a firmar* panel to add more files
3. Place the signature on the canvas (click **+ Agregar Firma**); optionally use **Replicar en todas las páginas**
4. Click **FIRMAR** — all documents are signed automatically and saved as `<name>_firmado.pdf` next to each original

> Clicking a filename in the *Documentos a firmar* list switches the canvas preview to that file.

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
git tag v2.1.1
git push origin main --tags
```

GitHub Actions will build and publish `FirmaRapida.exe` to the Releases page automatically.

## Security note

Your certificate password is stored **locally only**, encrypted with a machine-specific key using the `cryptography` library (Fernet). It is never transmitted anywhere.
