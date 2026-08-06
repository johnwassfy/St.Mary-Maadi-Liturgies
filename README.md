# St. Mary Maadi Liturgies

![St.Mary Maadi Liturgies Logo](Data/Logo.ico)

## Overview

St. Mary Maadi Liturgies is a desktop application created for the Coptic Orthodox Church of St. Mary in Maadi. The program generates ready-to-view PowerPoint presentations for prayers and liturgical services, taking into account the Coptic liturgical calendar, seasons, feast days, and special occasions.

The application includes a built-in Coptic calendar and algorithms to detect liturgical seasons and convert dates, so presentations are automatically tailored for the correct day and service.

## Key Features

- Dynamic liturgical content generated for the selected date and service
- Season-aware presentations (Great Lent, Holy Week, Pentecost, Kiahk, etc.)
- Built-in Coptic calendar with date conversion and season detection
- Support for multiple services:
  - Sunday Liturgy
  - Children's Liturgy
  - Matins (باكر)
  - Vespers (عشية)
  - Holy Week services
  - Special feast days and occasions
- Bible readings appropriate for each day
- Psalmodia support, including Kiahk psalmodia (الإبصلمودية)
- Hymns and praises library (المدائح)
- Special handling for episcopal (bishop) presence during services
- Automatic update checker and installer

## Installation

### Download

You can download the latest installer directly from the repository:

[Download StMaryMaadiInstaller.exe](https://github.com/johnwassfy/St.Mary-Maadi-Liturgies/raw/master/StMaryMaadiInstaller.exe)

### Installation Steps

1. Download the installer from the link above.
2. Right-click the downloaded file and choose "Run as administrator".
3. Follow the on-screen instructions to complete the installation.
4. Launch the application from the desktop shortcut or Start menu.

### System Requirements

- Windows 10 or later
- Microsoft PowerPoint 2016 or later (for opening generated presentations)
- 4 GB RAM minimum (8 GB recommended)
- 500 MB free disk space

## Usage

1. Launch the application. The main interface displays the current Coptic date and detected liturgical season.
2. Choose the service type from the menu (Liturgy, Matins, Vespers, etc.).
3. To change the date, click the date display and select the desired date; the application will update content accordingly.
4. Generate the presentation by selecting the service and pressing the generate/open button — a PowerPoint file will be created and opened automatically.
5. For special services (e.g., Holy Week, Lakan), use the "المناسبات" (Occasions) section to select the appropriate liturgy sequence.
6. If a bishop is present, enable "في حضور الأسقف" and optionally specify guest bishops to adjust the content and rubrics.

## Directory Structure

- Data/ — images, icons, templates, and other static resources
- Python files — main application logic and utilities
- Excel files — configuration, mappings, and liturgical data (read by the application)

## Technical Details

- Language: Python (PyQt5 for UI)
- PowerPoint integration via Win32 COM (win32com.client)
- Custom Coptic calendar algorithms for date conversion and season detection
- Uses asynchronous programming patterns to keep the UI responsive during generation and updates

## Updates and Maintenance

The application includes an automatic update system:

- Click "البحث عن تحديث" to check for updates manually.
- When updates are available, click "تحديث البرنامج" to download and install.

Release notes and detailed change history are shown inside the application when updates are available.

## Version History

- Current Version: 2.3.3 (as of June 5, 2025)
- See the in-app update notes for a history of changes and bug fixes.

If you maintain a formal changelog or use GitHub Releases, consider adding a Releases page with versioned notes.

## License

This software is developed for the specific use of St. Mary Coptic Orthodox Church — Maadi. If you want to reuse or redistribute the code, please contact the maintainer for licensing details.

## Contact & Support

For support, questions, or to report issues, contact the developer:

- Email: johnwassfy@gmail.com
- GitHub: https://github.com/johnwassfy/St.Mary-Maadi-Liturgies

---

© 2023-2025 St. Mary Coptic Orthodox Church - Maadi
