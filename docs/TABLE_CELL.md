# Table cell adaptive fitting

## What it means

When the caret is **inside a Word table cell** and you smart-paste an image, the image is sized against the **cell’s usable width/height**, not the full page content width.

Example:

- Page content width = 16 cm
- Current cell width = 6 cm
- Style max width = 92% of content

Without cell fitting → image aims near 14.7 cm and overflows the cell.  
With cell fitting → image aims near 5.5 cm and stays inside the cell.

## Why this exists

Word’s built-in inline paste often clamps to page/text margins. In tables, that can still feel wrong: screenshots become huge relative to the cell and force awkward wrapping or manual dragging.

## Preview

In the settings window, enable **Preview: simulate table cell** to see an amber dashed cell box and how the sample screenshot fits inside it when parameters change.
