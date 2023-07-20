# supporting languages
- [日本語](https://github.com/Hoshimikan6490/pptx_password_remover/blob/master/README_ja.md)
- **[English](https://github.com/Hoshimikan6490/pptx_password_remover/blob/master/README_en.md)** ←

# About this program
This programme removes passwords placed on PowerPoint files with the extension ".pptx".  

# CAUTION
This programme can be used to delete passwords against the intention of the PPTX file creator. This is not a good practice, so please ensure that the file with the password deleted is used only by the person who deleted it. Please refrain from distributing it or any other action.  
The author of this programme is not liable for any damage caused by the use of this programme. Examples of damages include disputes with the PPTX file creator or damage to the PPTX file after use of this programme.

# How to use
1. download this program or clone it using git.
2. open a terminal in the top directory in the programme you copied (the folder with this file).
3. Run the command `npm i`.
4. Execute the command `npx nodemon index.js` or `node index.js`. 
5. the URL will be displayed in the terminal. So, please access that url.
6. follow the on-screen instructions to select the PPTX file and press 'GO'.
7. the PPTX file with the password removed will be downloaded.

# Technical information
## About the action of this program
The action of the Programme is as follows.
1. Set up a local web server with the express package.
2. Ask the user to select a file and copy it to the working folder named "uploads".
3. Change the extension of PPTX files in the working folder to zip.
4. Unzip the zip file.
5. Remove text containing password information.
6. Compressed as a zip file.
7. Change extension to PPTX.
8. Make the user download it.
9. Remove pptx file in the working folder.

# Packages used and their purpose
- adm-zip: To compress and decompress zip files.
- express: To set up a web server.
- express-fileupload: On the web server, to accept files.
- fs: To edit with files that have been uploaded.
- nodemon: To restart immediately when there are changes related to the web server.
- util: To implement the 'wait a few seconds' bug countermeasure.