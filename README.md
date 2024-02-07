# XLSX Image to Filename

## Instructions

1. Open _Windows Terminal_ app (Start menu -> search "Terminal")
2. Type the following command to see documentation:
   ```sh
   xlsx-image-to-filename --help
   ```

## Example Usage

```sh
xlsx-image-to-filename --file-path "oss图文对照表1.xlsx" --prefix "https://clearlyloc.sharepoint.com/sites/ProjectScreenshots/oss1/"
```

## Building from Source

To build from source:

```sh
deno compile -A --output ./xlsx-image-to-filename --target x86_64-pc-windows-msvc ./src/main.ts --__is-compiled
```
