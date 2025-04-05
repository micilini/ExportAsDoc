# ExportAsDoc

## Overview
`ExportAsDoc.exe` is a Python-based script converted into a Windows executable. Its purpose is to process JSON data and generate a formatted Word (.docx) document as output. This script is intended to be invoked by other applications rather than executed directly by the user.

`ExportAsDoc` was created as a module for the [Minimal Text Editor (Lite)](https://github.com/micilini/MinimalTextEditorLite) application.

---

## Table of Contents
1. [How to Generate the Executable](#how-to-generate-the-executable)
2. [How to Use the Script](#how-to-use-the-script)
3. [Example Usage](#example-usage)

---

## How to Generate the Executable

### Requirements
- Python 3.8 or higher
- `pip` package manager
- The following Python packages:
  - `docx`
  - `pillow`
  - `pyinstaller`
  - `json`

### Steps
1. **Install Required Packages**:
   Run the following command to install the dependencies:
   ```bash
   pip install python-docx pillow pyinstaller
   ```

2. **Create the Executable**:
   Use `pyinstaller` to package the script into an executable:
   ```bash
   pyinstaller --onefile --name=ExportAsDoc ExportAsDoc.py
   ```
   - The `--onefile` flag ensures the executable is a single file.
   - The `--name` flag specifies the output executable's name.

   If the code below don't work, try to create the executable with PIL support:

   ```bash
   pyinstaller --onefile --hidden-import=PIL --hidden-import=PIL._imaging --hidden-import=PIL.Image ExportAsDoc.py
   ```

3. **Locate the Executable**:
   After running the above command, the `ExportAsDoc.exe` file will be available in the `dist` folder.

---

## How to Use the Script

### Important Notes
- This script is designed to be called programmatically by other applications. It should not be opened directly by the user.
- The script accepts a single argument: the file path of the JSON input.

### Workflow
1. Save the JSON data to a local file.
2. Pass the file path as an argument when invoking the executable.
3. Capture the generated `.docx` file output from the standard output.

### Executing the script with terminal
You can execute the executable with terminal (cmd, git bash and others). Inside the ```dist``` folder, you will find ```ExportAsDoc.exe``` with ```note.json```, navigate with your terminal to that folder, and then run the following command:

```bash
./ExportAsDoc.exe note.json > result.docx
```

Finally, a file named ```result.docx``` will be created in the same folder.

---

## Example Usage

### Using the Executable in C#
Below is an example of how to call `ExportAsDoc.exe` from a C# application:

```csharp
// Save JSON data to a temporary file
string tempJsonFilePath = Path.Combine(Path.GetTempPath(), "data.json");
File.WriteAllText(tempJsonFilePath, jsonData);

// Execute the process asynchronously
var result = await Task.Run(() =>
{
    var processStartInfo = new ProcessStartInfo
    {
        FileName = "Modules\\Export\\ExportAsDoc.exe",
        Arguments = $"\"{tempJsonFilePath}\"",
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false,
        CreateNoWindow = true
    };

    using (var process = new Process { StartInfo = processStartInfo })
    {
        process.Start();

        using (var memoryStream = new MemoryStream())
        {
            process.StandardOutput.BaseStream.CopyTo(memoryStream);
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                var error = process.StandardError.ReadToEnd();
                throw new Exception(error);
            }

            return memoryStream.ToArray(); // Returns the binary DOCX file
        }
    }
});
```

### Key Points
- The JSON file is saved temporarily before being passed to the executable.
- The executable's standard output is captured to retrieve the binary content of the generated `.docx` file.
- The process is run in a separate shell with no visible window to provide a seamless user experience.

---

## Contributing

Feel free to open issues or submit pull requests to improve the script.

---

## License

This script is open-source and available under the MIT License.