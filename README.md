# SOFIA - System för Operativ Förvaltning, Incidentöversikt och Analys

SOFIA is a simple web-based system for keeping track the status of various systems without an external database; SOFIA uses a xlsx file.
Swedish front end, should be easy to adapt for your needs. 

![image](https://github.com/user-attachments/assets/987b6e81-6af3-4f10-b710-b05c3c70c46e)


## Features

- Track the status of systems (Tested OK, Not Tested).
- Add and remove systems from the list.
- Clear the status of all systems.
- View statistics on the number of systems tested.

## Installation

### Prerequisites

- PHP (7.4 or higher)
- Composer
- Web server (e.g., Lighttpd, Nginx)
- PhpSpreadsheet library

### Steps

1.  **Clone the Repository**:

   git clone https://github.com/FilleMang/SOFIA.git
   cd sofia

2.  **Install Dependencies:**

composer install

3.  **Configure the Web Server:**

Set up your web server to point to the sofia directory.
Ensure that the web server has read and write permissions for the directory.

4.  **Prepare the Excel File:**

Create a source.xlsx file in the root directory or use the example provided.

5.  **Access the Application:**

Open your web browser and navigate to the URL where you have set up the application.
