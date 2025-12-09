# MySheets

[![.NET](https://img.shields.io/badge/.NET-9.0-blue.svg)](https://dotnet.microsoft.com/download)
[![C#](https://img.shields.io/badge/C%23-12-239120.svg)](https://docs.microsoft.com/en-us/dotnet/csharp/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

A lightweight, cross-platform spreadsheet application built with .NET 9.0. Designed for performance using sparse memory structures and a custom-built formula compiler.

##   Features

- **Modern UI:** Clean and intuitive interface with a dedicated formula bar for easy editing (supports visual area selection, click-to-link cell references etc..).
- **Smart Formulas:** Supports complex math expressions and cell references (e.g., `=A1+B2`, `=SUM(A1:H7)`).
- **File Support:** Save and load your spreadsheets in **JSON** and **Excel (.xlsx)** formats.
- **Instant Updates:** Changing a value automatically recalculates all related cells in real-time.
- **High Performance:** Optimized to run fast and use minimal memory.
- **Cross-Platform:** Works natively on Windows, macOS, and Linux.

## Tech Stack

- **Core:** C# 12, .NET 9.0
- **Architecture:** Clean Architecture (Core logic decoupled from UI)
- **UI Framework:** Avalonia UI (MVVM)
- **Testing:** xUnit

## Getting Started

```bash
git clone https://github.com/zahorodnyi/my-sheets.git
cd my-sheets/MySheets.UI
dotnet run
```