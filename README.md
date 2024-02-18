# XLSX Image to Filename

## Instructions (Windows)

1. Save `xlsx-image-to-filename.exe` to your local file system. In this example, we assume the following folder structure:
   ```
   .
   ├── xlsx-image-to-filename.exe
   └── input_files/
      ├── input-file-1.xlsx
      ├── input-file-2.xlsx
      └── input-file-3.xlsx
   ```
2. Open Windows Explorer.
3. Right-click within the root folder, then select _Open in Terminal_
   **Note**: Alternatively, you can open _Windows Terminal_ app (Start menu -> search "Terminal"), then navigate to the folder using `cd`.
4. Type the following command to see documentation:
   ```sh
   xlsx-image-to-filename --help
   ```

### Example usage

In our example folder structure above, we could run the following to convert all files in the `input_files` subfolder, writing the output files to the `output_files\{{DATE}}` subfolder (where `{{DATE}}` is today's date), and using the prefix `https://clearlyloc.sharepoint.com/sites/SiteName/path/to/library/`:

```sh
xlsx-image-to-filename input_files --out-path output_files\{{DATE}} --prefix https://clearlyloc.sharepoint.com/sites/SiteName/path/to/library
```

The folder structure afterwards would look something like this:

```
.
├── xlsx-image-to-filename.exe
├── input_files/
│   ├── input-file-1.xlsx
│   ├── input-file-2.xlsx
│   └── input-file-3.xlsx
└── output_files/
    └── 20240218/
        ├── output-file-1.xlsx
        ├── output-file-2.xlsx
        └── output-file-3.xlsx
```
