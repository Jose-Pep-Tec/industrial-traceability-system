# Production Control System

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

### Python Version
**CRITICAL**: This application requires **Python 3.5 ONLY**.

The code does NOT work with Python 3.6 or higher. It does not use f-strings or other Python 3.6+ features.

### Operating System
- Windows 7 or later (for Zebra printer support)

### Dependencies

| Library | Version | Purpose |
|---------|---------|---------|
| Python | 3.5 ONLY | Core interpreter |
| Pillow | 5.0.0 | Image processing for label preview |
| python-barcode | 0.13.0 | Barcode generation |
| pywin32 | 223 | Windows printer API |

## Installation

### 1. Python Version

**IMPORTANT**: Make sure you have Python 3.5 installed:

```bash
python --version
# Must output: Python 3.5.x
