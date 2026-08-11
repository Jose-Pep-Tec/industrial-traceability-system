# industrial-traceability-system

A comprehensive desktop application for manufacturing production management, order tracking, and label printing.

## Overview

This system provides a complete solution for managing production workflows, including order creation, production registration, label printing, and reporting. Built with Python 3.5 and Tkinter, it offers a user-friendly interface for shop floor operators and production managers.

## Features

### Order Management
- Create new orders with client, part number, destination, and target quantity
- Work on active orders with real-time progress tracking
- Edit existing orders (date, destination, target quantity)
- Cancel or reactivate orders
- Hide completed orders (soft delete)

### Production Registration
- Register production against active orders
- Manual registration without order association
- Automatic serial number generation (sequential per order)
- Remainder detection and loading from previous orders
- Excess production adjustment with automatic remainder calculation
- Duplicate detection to prevent data entry errors

### Label Printing
- Generate ZPL code for Zebra printers
- Print individual labels or entire batches
- Preview labels before printing
- Support for partial bags (last bag with different quantity)
- Re-print previously printed labels

### Reporting
- Generate CSV reports with filters (date range, client, part number)
- Export data for external analysis
- Order-specific reports

### Administration
- Manage clients and their destinations
- Manage parts with client association and complement configuration
- Configure label printer selection

### Data Management
- Built-in SQLite database
- Record editor for advanced users (password protected)
- Export data to CSV

## Requirements

### System Requirements
- **Operating System**: Windows 7 or later (for Zebra printer support)
- **Python**: Version 3.5 or higher

### Python Dependencies

| Library | Version | Purpose |
|---------|---------|---------|
| tkinter | Built-in | GUI framework |
| sqlite3 | Built-in | Database management |
| json | Built-in | Configuration file handling |
| csv | Built-in | Report export |
| datetime | Built-in | Date and time handling |
| os | Built-in | File system operations |
| re | Built-in | Regular expressions |
| PIL (Pillow) | 5.0.0+ | Image processing for label preview |
| python-barcode | 0.13.0+ | Barcode generation |
| pywin32 | 223+ | Windows printer API |

### Hardware Requirements
- **Minimum RAM**: 512 MB
- **Disk Space**: 50 MB
- **Display**: 1024 x 768 resolution
- **Printer**: Zebra compatible label printer (ZPL support)

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/production-control-system.git
cd production-control-system
