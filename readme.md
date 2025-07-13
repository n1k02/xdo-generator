## 🚀 XDO Generator

Easily create your XDO executable with a single command!

### How to Generate the Executable

1. Open your terminal.
2. Run the following command:

    ```sh
    pyinstaller --onefile --name xdo_generator src/main.py
    ```

After the process completes, you'll find your `xdo_generator.exe` in the `dist` folder.


### Usage

1. Place `xdo_generator.exe` in the root directory of your project.
2. Add your SQL model code to text files named `G1.txt`, `G2.txt`, etc.
3. Rename your Excel file to `template.xlsx`.
4. Run `xdo_generator.exe` to generate the xls template.

> Alternatively, you can run `main.py` directly with Python.  
> Make sure to install all required dependencies first.

---

