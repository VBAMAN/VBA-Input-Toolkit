# VBA-Input-Toolkit
A lightweight worksheet-based input toolkit for Excel VBA.

VBA Input Toolkit is an open-source toolkit that brings worksheet-based input assistance to Microsoft Excel.

It provides a movable input palette, automatic numbering, default value insertion, and other utilities to improve repetitive data entry without relying on UserForms.

It is designed for repetitive worksheet data entry where speed and simplicity are more important than complex UserForms.

Instead of opening dialog windows, users can select items directly from a movable worksheet palette. This approach keeps the worksheet visible at all times and reduces the number of keyboard and mouse operations required for data entry.

The toolkit is lightweight, easy to customize, and intended to be used as a foundation for Excel VBA input solutions.

## Features

* Movable worksheet input palette
* Worksheet-based input without UserForms
* Automatic serial numbering
* Optional default quantity input
* Automatic palette follow mode
* Overwrite confirmation
* Lightweight and easy to customize
* Pure VBA implementation (no external libraries)

## Requirements

* Microsoft Excel with VBA support
* Windows environment
* Macro-enabled workbook (`.xlsm`)

Developed and tested on Microsoft Excel 2007.

The toolkit uses standard VBA features only and does not require any external libraries or dependencies.

## Getting Started

1. Open the sample workbook in Microsoft Excel.
2. Enable VBA macros when prompted.
3. Click **ON** to enable palette input mode.
4. Select an item from the worksheet palette.
5. Click a target cell in the input column to insert the selected value.
6. Continue entering data. The palette can automatically follow the current editing position if enabled.

To stop palette input, click **OFF**.

## Configuration

The toolkit can be customized by editing the configuration constants in the `M01_Core` module.

Typical settings include:

* Palette position
* Palette follow behavior
* Automatic quantity input
* Follow interval
* Input column
* Overwrite confirmation

The default settings are intended for the included sample workbook but can be easily adapted to different worksheet layouts.

## Project Structure

| Module                 | Description                                                       |
| ---------------------- | ----------------------------------------------------------------- |
| `M01_Core`             | Global configuration, constants, and runtime state.               |
| `M02_PaletteEngine`    | Palette movement, positioning, and follow behavior.               |
| `M03_InputEngine`      | Palette input processing and worksheet data entry.                |
| `M04_AutoNumberEngine` | Automatic serial number generation.                               |
| `M05_Utilities`        | Shared helper functions used by the toolkit.                      |
| `Worksheet`            | Event handlers that connect the worksheet to the toolkit engines. |

The toolkit is organized into independent modules to simplify maintenance, customization, and future expansion.

## License

This project is released under the MIT License.

You are free to use, modify, and distribute this software in accordance with the terms of the license. See the `LICENSE` file for details.

## Author

Developed by VBAMAN


