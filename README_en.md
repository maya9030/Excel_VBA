# VBA Tools Collection

This repository is a collection of **VBA macros for Excel** that I created as personal projects. Each tool is organized into its own subdirectory with corresponding source code and documentation.

---

## Contents

### 1. `ResetSheet/`

A macro that **resets the currently active worksheet** to a clean state. Useful for clearing content, removing borders, and deleting images before reusing templates or starting fresh.

**Key Features:**
- Prompts user for confirmation before action
- Deletes all images (shapes of type `msoPicture`)
- Clears all cell contents and borders
- Displays a confirmation message when done

Use case: **resetting a worksheet to default state before re-entry**

---

### 2. `SplitNumbers/`

A macro built for **memory athletes and trainers**. It takes a long numeric string, splits it into 4-digit chunks, and converts each chunk into two 2-digit numbers. These are then used to **retrieve associated words** from a 10×10 matrix using `VLOOKUP`.

**Key Features:**
- Supports numeric chunking and formatting
- Dynamically builds VLOOKUP formulas for matrix-based word lookup
- Interactive cell selection for flexible input/output
- Ideal for memorization training using number→image conversion

Use case: **generating word associations from digit strings for memory training**

---

## How to Use

1. Open the desired subdirectory (`ResetSheet` or `SplitNumbers`)
2. Review the included `README.md` for setup and usage instructions
3. Copy the code into your Excel VBA editor and run the macro

---

## Requirements

- Microsoft Excel (VBA-enabled)
- Basic familiarity with running macros
- (For `SplitNumbers`) A 10×10 matrix in `B2:L12` containing word associations

---

## License

MIT License

---

Feel free to fork, adapt, and expand these macros to suit your workflow.
