# Supporting Languages
- [日本語](./README_ja.md)
- **[English](./README_en.md)** ←

# About
This local web tool supports two operations for PowerPoint files:

- Remove password lock from PPTX/PPSX
- Convert PPSX to PPTX

# Caution
Password removal can conflict with the original creator's intent. Please use this tool only for files you are allowed to handle.  
The author is not responsible for any damage caused by using this program.

# Usage
1. Clone or download this repository.
2. Open a terminal in the project root.
3. Run `npm i`.
4. Run `node index.js` or `npx nodemon index.js`.
5. Open the URL shown in the terminal (usually http://localhost:6490).
6. Choose an operation in the UI, upload a file, and click GO.
7. Download the processed file.

# API
The UI sends one of these POST requests:

- Password removal: `/api/progress?type=removePassword`
- PPSX to PPTX: `/api/progress?type=convertToPPTX`

Form field:

- `file` (required): `.pptx` or `.ppsx`
- In `convertToPPTX` mode, only `.ppsx` is accepted

# Major Error Responses
- `No files were uploaded.`
- `Only .pptx or .ppsx files are supported.`
- `convertToPPTX mode only supports .ppsx files.`
- `Uploaded PowerPoint file does not appear to be password-protected.`
	- Returned in removePassword mode when no `p:modifyVerifier` tag exists

# Technical Overview
1. Start a local server with `express`
2. Save uploaded file temporarily
3. Read internal XML with `jszip`
4. Rewrite XML depending on selected mode
5. Rebuild file and return it as download
6. Delete temporary file unless debug mode is enabled

# Packages
- `express`
- `express-fileupload`
- `jszip`
- `nodemon` (development)
