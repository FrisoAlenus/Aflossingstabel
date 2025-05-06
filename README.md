# Aflossingstabel

A Python-based GUI tool to help fill in an "Aflossingstabel" (Amortization Table) Excel template, adjust the number of interest periods, and save the modified file while preserving formulas.

## Features

* Fills predefined cells in an Excel template with user input.
* Dynamically adjusts the number of visible interest periods by deleting excess rows.
* Handles different starting rows for period calculations based on loan type (e.g., "Mezzanine").
* Preserves existing formulas in the Excel template.
* Allows selection of "Kredietverstrekker" from a dropdown (LRM, Mijnen).
* Defaults "Begindatum" (Start Date) to the current date.
* User-friendly GUI built with Tkinter.

## Requirements

* Python 3.7+
* openpyxl library
