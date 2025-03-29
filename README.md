# PowerPoint VBA Script for Processing Prize Categories

## Overview

This VBA script automates the process of generating and updating PowerPoint slides based on prize categories. It reads data from CSV files and updates slide content dynamically.

## Features

- Reads prize category data from CSV files.
- Updates slides dynamically with text content.
- Handles missing files gracefully by prompting the user.
- Deletes existing slides before inserting new ones.

## Prerequisites

- Microsoft PowerPoint with VBA enabled.
- CSV files containing prize category data.
- Folder structure: The script expects CSV files inside a `files` folder within the same directory as the PowerPoint file.

## File Structure

```txt
YourPresentation.pptm
files/
    top3.csv
    top10.csv
    gold.csv
    silver.csv
    bronze.csv
    honorable.csv
```

## Script Breakdown

### `Sub Main()`

Calls `ProcessPrizeCategory` for each prize category with its corresponding slide ID.

### `Sub ProcessPrizeCategory(id As Long, prize As String)`

- Reads the corresponding CSV file.
- Splits text content and processes each entry.
- Calls `GenerateSlide` to create slides.
- Calls `DeleteSlide` to remove previous slides.

### `Sub GenerateSlide(id As Long, text As Variant)`

Duplicates an existing slide and updates the text in the shape named `names`.

### `Sub DeleteSlide(id As Long)`

Deletes the slide with the given ID.

### `Function TextFileToArray(filePath As String, Optional LineSeparator As String = vbLf) As Variant`

Reads a CSV file and returns an array of lines.

### `Function GetSplitArray(SourceArray As Variant, Optional ColumnDelimiter As String = ",") As Variant`

Splits lines into a 2D array based on a column delimiter.

## How to Use

1. Open the PowerPoint file and enable macros.
2. Ensure the required CSV files are placed inside the `files` folder.
3. Run the `Main` macro from the VBA editor.

## Error Handling

- If a required CSV file is missing, a message box prompts the user to create an empty file.
