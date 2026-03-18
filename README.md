# Excel VBA Taxonomy Macro

This project contains a simple Excel VBA macro for working with taxonomy topic selections and text concatenation.

## Overview

The VBA code includes:

- a custom function to combine values from multiple cells into a single text string
- a macro that checks whether the active cell is in the taxonomy column
- a user form trigger for taxonomy topic selection

## Features

### `CONCATENATEMULTIPLE`
This user-defined function concatenates values from a selected range of cells using a chosen separator.

Example use in Excel:

```excel
=CONCATENATEMULTIPLE(A1:A5, ", ")
